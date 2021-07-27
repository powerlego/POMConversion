package org.conversion.tasks;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.*;
import org.conversion.sql.DBConnectionPool;
import org.conversion.tasks.assemblies.AssemblyItem;
import org.conversion.tasks.assemblies.AssemblyMaster;
import org.conversion.utils.Utils;

import javax.sql.rowset.CachedRowSet;
import javax.sql.rowset.RowSetFactory;
import javax.sql.rowset.RowSetProvider;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Paths;
import java.sql.Connection;
import java.sql.SQLException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * @author Nicholas Curl
 */
public class Packages implements Task {

    /**
     * The instance of the logger
     */
    private static final Logger logger = LogManager.getLogger(Packages.class);

    @Override
    public void run() {
        try {
            Connection connection = DBConnectionPool.getConnection();
            RowSetFactory factory = RowSetProvider.newFactory();
            CachedRowSet result = factory.createCachedRowSet();
            result.setCommand(Utils.joinSQLLine("SELECT [KEY],\n" +
                                                "       ItemFile.NUM,\n" +
                                                "       ItemKey,\n" +
                                                "       Quantity,\n" +
                                                "       (SELECT Name FROM ItemFile WHERE [KEY] = ItemKey) AS Name,\n" +
                                                "       (SELECT NUM FROM ItemFile WHERE [KEY] = ItemKey) AS ItemNum\n" +
                                                "FROM ItemFile\n" +
                                                "         LEFT OUTER JOIN ItemKits ON ItemKits.Num = ItemFile.NUM\n" +
                                                "WHERE [KEY] LIKE '%PKG'\n" +
                                                "  AND [KEY] NOT LIKE 'HD-%'\n" +
                                                "  AND ItemKey IS NOT NULL"));
            result.execute(connection);
            Map<AssemblyMaster, List<AssemblyItem>> assemblyItemMap = new HashMap<>();
            while (result.next()) {
                String assemblyItemKey = result.getString(3);
                int assemblyItemQty;
                if (result.getObject(4) instanceof Double) {
                    assemblyItemQty = (int) result.getDouble(4);
                }
                else if (result.getObject(4) instanceof Integer) {
                    assemblyItemQty = result.getInt(4);
                }
                else {
                    assemblyItemQty = Integer.MIN_VALUE;
                }
                if (assemblyItemQty == Integer.MIN_VALUE) {
                    logger.warn("Invalid Quantity for Item {}", assemblyItemKey);
                    continue;
                }
                String assemblyKey = result.getString(1);
                int assemblyItemNum = result.getInt(2);
                AssemblyMaster master = new AssemblyMaster(assemblyKey, assemblyItemNum);
                if (!assemblyItemMap.containsKey(master)) {
                    assemblyItemMap.put(master, new ArrayList<>());
                }
                String assemblyItemDesc = result.getString(5);
                int assemblyItemItemNum = result.getInt(6);
                AssemblyItem assemblyItem = new AssemblyItem(assemblyItemItemNum,
                                                             assemblyItemKey,
                                                             assemblyItemDesc,
                                                             assemblyItemQty
                );
                assemblyItemMap.get(master).add(assemblyItem);
            }
            connection.close();
            XSSFWorkbook workbook = new XSSFWorkbook();
            XSSFSheet sheet = workbook.createSheet("NS BOM Import");
            XSSFRow headerRow = sheet.createRow(0);
            XSSFCell headerCell = headerRow.createCell(0);
            XSSFCellStyle headerStyle = workbook.createCellStyle();
            headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            headerStyle.setFillForegroundColor(new XSSFColor(IndexedColors.GREY_25_PERCENT,
                                                             new DefaultIndexedColorMap()
            ));
            headerStyle.setFillBackgroundColor((short) 64);
            XSSFFont font = workbook.findFont(true,
                                              IndexedColors.AUTOMATIC.getIndex(),
                                              XSSFFont.DEFAULT_FONT_SIZE,
                                              XSSFFont.DEFAULT_FONT_NAME,
                                              false,
                                              false,
                                              XSSFFont.SS_NONE,
                                              XSSFFont.U_NONE
            );
            if (font == null) {
                font = workbook.createFont();
                font.setBold(true);
                font.setColor(IndexedColors.AUTOMATIC.getIndex());
                font.setFontHeightInPoints(XSSFFont.DEFAULT_FONT_SIZE);
                font.setFontName(XSSFFont.DEFAULT_FONT_NAME);
                font.setItalic(false);
                font.setStrikeout(false);
                font.setTypeOffset(XSSFFont.SS_NONE);
                font.setUnderline(XSSFFont.U_NONE);
            }
            headerStyle.setFont(font);
            headerCell.setCellValue("Name");
            headerCell.setCellStyle(headerStyle);
            int rowCount = 1;
            for (AssemblyMaster assemblyMaster : assemblyItemMap.keySet()) {
                XSSFRow row = sheet.createRow(rowCount);
                XSSFCell cell = row.createCell(0);
                cell.setCellValue(assemblyMaster.getKey());
                rowCount++;
            }
            sheet.autoSizeColumn(0);
            sheet = workbook.createSheet("NS BOM Revision Import");
            headerRow = sheet.createRow(0);
            int itemCount = 0;
            for (List<AssemblyItem> assemblyItems : assemblyItemMap.values()) {
                if (assemblyItems.size() > itemCount) {
                    itemCount = assemblyItems.size();
                }
            }
            headerCell = headerRow.createCell(0);
            headerCell.setCellValue("Item Number");
            headerCell.setCellStyle(headerStyle);
            headerCell = headerRow.createCell(1);
            headerCell.setCellValue("Bill of Materials");
            headerCell.setCellStyle(headerStyle);
            headerCell = headerRow.createCell(2);
            headerCell.setCellValue("Name");
            headerCell.setCellStyle(headerStyle);
            int cellCount = 3;

            for (int i = 0; i < itemCount; i++) {
                XSSFCell cell = headerRow.createCell(cellCount);
                cell.setCellValue("CI Item Number " + (i + 1));
                cell.setCellStyle(headerStyle);
                cellCount++;
                cell = headerRow.createCell(cellCount);
                cell.setCellValue("CI Item " + (i + 1));
                cell.setCellStyle(headerStyle);
                cellCount++;
                cell = headerRow.createCell(cellCount);
                cell.setCellValue("CI Quantity " + (i + 1));
                cell.setCellStyle(headerStyle);
                cellCount++;
            }
            rowCount = 1;
            for (AssemblyMaster assemblyMaster : assemblyItemMap.keySet()) {
                XSSFRow row = sheet.createRow(rowCount);
                XSSFCell cell = row.createCell(0);
                cell.setCellValue(assemblyMaster.getItemNum());
                cell = row.createCell(1);
                cell.setCellValue(assemblyMaster.getKey());
                cell = row.createCell(2);
                cell.setCellValue(assemblyMaster.getKey() + " Rev A");
                rowCount++;
                cellCount = 3;
                List<AssemblyItem> assemblyItems = assemblyItemMap.get(assemblyMaster);
                for (int i = 0; i < itemCount; i++) {
                    try {
                        AssemblyItem assemblyItem = assemblyItems.get(i);
                        cell = row.createCell(cellCount);
                        cell.setCellValue(assemblyItem.getItemNum());
                        cellCount++;
                        cell = row.createCell(cellCount);
                        cell.setCellValue(assemblyItem.getKey() + " " + assemblyItem.getDescription());
                        cellCount++;
                        cell = row.createCell(cellCount);
                        cell.setCellValue(assemblyItem.getQty());
                        cellCount++;
                    }
                    catch (IndexOutOfBoundsException e) {
                        for (int j = 0; j < 3; j++) {
                            cell = row.createCell(cellCount);
                            cellCount++;
                            cell.setBlank();
                        }
                    }
                }
            }
            for (int i = 0; i < Utils.getLastColumn(sheet); i++) {
                sheet.autoSizeColumn(i);
            }
            try {
                FileOutputStream fileOutputStream = new FileOutputStream(Paths.get("./package_items_bom.xlsx")
                                                                              .toFile());
                workbook.write(fileOutputStream);
                workbook.close();
                fileOutputStream.close();
            }
            catch (IOException e) {
                logger.fatal("Unable to write workbook", e);
                System.exit(1);
            }
        }
        catch (SQLException throwables) {
            logger.fatal(throwables.getMessage(), throwables);
            System.exit(1);
        }
    }

}

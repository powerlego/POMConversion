package org.conversion.tasks;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.conversion.tasks.assemblies.AssemblyItem;
import org.conversion.tasks.assemblies.AssemblyMaster;
import org.conversion.tasks.assemblies.ItemGroup;
import org.conversion.utils.Utils;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Paths;
import java.util.*;

/**
 * @author Nicholas Curl
 */
public class Configurators implements Task {

    /**
     * The instance of the logger
     */
    private static final Logger logger = LogManager.getLogger(Configurators.class);

    @Override
    public void run() {
        File[] configurators = Paths.get("./configurators").toFile().listFiles();
        if (configurators != null && configurators.length > 0) {
            for (File configurator : configurators) {
                Map<Integer, AssemblyMaster> boms = new TreeMap<>();
                XSSFWorkbook inputWorkbook;
                OPCPackage opcPackage;
                try {
                    opcPackage = OPCPackage.open(configurator);
                    inputWorkbook = new XSSFWorkbook(opcPackage);
                }
                catch (InvalidFormatException | IOException exception) {
                    inputWorkbook = null;
                    opcPackage = null;
                    logger.fatal("Unable to read workbook", exception);
                    System.exit(1);
                }
                XSSFWorkbook workbook = new XSSFWorkbook();
                for (Sheet sheet : inputWorkbook) {
                    XSSFSheet xssfSheet = (XSSFSheet) sheet;
                    if (!xssfSheet.getSheetName().toLowerCase(Locale.ROOT).contains("master")) {
                        for (Row row : xssfSheet) {
                            XSSFRow xssfRow = (XSSFRow) row;
                            XSSFCell xssfCell = xssfRow.getCell(0);
                            if (xssfCell == null || xssfCell.getCellType().equals(CellType.BLANK)) {
                                Object value = Utils.getCellValue(xssfRow.getCell(1),
                                                                  inputWorkbook.getCreationHelper()
                                                                               .createFormulaEvaluator()
                                );
                                int itemNum;
                                if (value instanceof Integer) {
                                    itemNum = (int) value;
                                }
                                else if (value != null) {
                                    try {
                                        itemNum = (int) Double.parseDouble(value.toString());
                                    }
                                    catch (NumberFormatException e) {
                                        itemNum = 0;
                                        logger.fatal("Unable to get Item Number", e);
                                        System.exit(1);
                                    }
                                }
                                else {
                                    continue;
                                }
                                Object nameValue = Utils.getCellValue(xssfRow.getCell(2),
                                                                      inputWorkbook.getCreationHelper()
                                                                                   .createFormulaEvaluator()
                                );
                                String name;
                                if (nameValue != null) {
                                    name = nameValue.toString();
                                }
                                else {
                                    name = "";
                                    logger.fatal("Unable to get item name");
                                    System.exit(1);
                                }
                                boms.put(itemNum, new AssemblyMaster(name, itemNum));
                            }
                            else {
                                break;
                            }
                        }
                    }
                }
                XSSFSheet masterSheet = null;
                for (Sheet sheet : inputWorkbook) {
                    if (sheet.getSheetName().toLowerCase(Locale.ROOT).contains("master")) {
                        masterSheet = inputWorkbook.getSheet(sheet.getSheetName());
                    }
                }
                if (masterSheet != null) {
                    int lastCol = Utils.getLastColumn(masterSheet);
                    List<List<Object>> master = new ArrayList<>();
                    for (int i = 0; i < masterSheet.getLastRowNum() + 1; i++) {
                        XSSFRow sheetRow = masterSheet.getRow(i);
                        List<Object> row = new ArrayList<>();
                        for (int j = 0; j < lastCol; j++) {
                            XSSFCell cell = sheetRow.getCell(j);
                            Object value = Utils.getCellValue(cell,
                                                              inputWorkbook.getCreationHelper().createFormulaEvaluator()
                            );
                            row.add(value);
                        }
                        master.add(row);
                    }
                    XSSFSheet ns_bom_import = workbook.createSheet("NS BOM Import");
                    XSSFRow headerRow = ns_bom_import.createRow(0);
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
                    for (AssemblyMaster assemblyMaster : boms.values()) {
                        XSSFRow row = ns_bom_import.createRow(rowCount);
                        XSSFCell cell = row.createCell(0);
                        cell.setCellValue(assemblyMaster.getKey());
                        rowCount++;
                    }
                    ns_bom_import.autoSizeColumn(0);
                    ns_bom_import = workbook.createSheet("NS BOM Revision Import");
                    headerRow = ns_bom_import.createRow(0);
                    int itemCount = 0;
                    int itemStart = 0;
                    for (List<Object> row : master) {
                        Object value = row.get(1);
                        if (value != null) {
                            if (value.toString().equalsIgnoreCase("Item Key")) {
                                itemStart = master.indexOf(row) + 1;
                                break;
                            }
                        }
                    }
                    Map<AssemblyMaster, List<AssemblyItem>> assemblies = new HashMap<>();
                    for (int itemNum : boms.keySet()) {
                        int col = Utils.findBOMColumn(itemNum, master.get(0));
                        AssemblyMaster assemblyMaster = boms.get(itemNum);
                        if (!assemblies.containsKey(assemblyMaster)) {
                            assemblies.put(assemblyMaster, new ArrayList<>());
                        }
                        if (col != Integer.MAX_VALUE) {
                            for (int i = itemStart; i < master.size(); i++) {
                                List<Object> masterRow = master.get(i);
                                Object quantityObject = masterRow.get(col);
                                if (quantityObject != null) {
                                    if (!quantityObject.toString().isBlank()) {
                                        int assemblyItemItemNum = getIntValue(masterRow.get(3));
                                        String assemblyItemKey = masterRow.get(1).toString();
                                        String assemblyItemDesc = masterRow.get(4).toString();
                                        int qty = getIntValue(quantityObject);
                                        AssemblyItem assemblyItem = new AssemblyItem(assemblyItemItemNum,
                                                                                     assemblyItemKey,
                                                                                     assemblyItemDesc,
                                                                                     qty
                                        );
                                        assemblies.get(assemblyMaster).add(assemblyItem);
                                    }
                                }
                            }
                        }
                        else {
                            logger.fatal("Unable to find column of BOM Item");
                            System.exit(1);
                        }
                    }
                    for (List<AssemblyItem> assemblyItems : assemblies.values()) {
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
                    for (AssemblyMaster assemblyMaster : assemblies.keySet()) {
                        XSSFRow row = ns_bom_import.createRow(rowCount);
                        XSSFCell cell = row.createCell(0);
                        cell.setCellValue(assemblyMaster.getItemNum());
                        cell = row.createCell(1);
                        cell.setCellValue(assemblyMaster.getKey());
                        cell = row.createCell(2);
                        cell.setCellValue(assemblyMaster.getKey() + " Rev A");
                        rowCount++;
                        int col = Utils.findBOMColumn(assemblyMaster.getItemNum(), master.get(0));
                        if (col != Integer.MAX_VALUE) {
                            cellCount = 3;
                            for (AssemblyItem assemblyItem : assemblies.get(assemblyMaster)) {
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
                        }
                        else {
                            logger.fatal("Unable to find column of BOM Item");
                            System.exit(1);
                        }
                    }
                    for (int i = 0; i < Utils.getLastColumn(ns_bom_import); i++) {
                        ns_bom_import.autoSizeColumn(i);
                    }
                    List<ItemGroup> itemGroups = new ArrayList<>();
                    for (Sheet sheet : inputWorkbook) {
                        XSSFSheet xssfSheet = (XSSFSheet) sheet;
                        if (!xssfSheet.getSheetName().toLowerCase(Locale.ROOT).contains("master")) {
                            XSSFRow headRow = xssfSheet.getRow(0);
                            for (Cell cell : headRow) {
                                if (cell == null ||
                                    Utils.getCellValue((XSSFCell) cell,
                                                       inputWorkbook.getCreationHelper().createFormulaEvaluator()
                                    ) ==
                                    null || Utils.getCellValue((XSSFCell) cell,
                                                               inputWorkbook.getCreationHelper()
                                                                            .createFormulaEvaluator()
                                ).toString().equalsIgnoreCase("Key") ||
                                    cell.getCellType() == CellType.BLANK) {
                                    continue;
                                }
                                int col = cell.getColumnIndex();
                                ItemGroup itemGroup = new ItemGroup(cell.getStringCellValue(),
                                                                    xssfSheet.getRow(2)
                                                                             .getCell(col)
                                                                             .getStringCellValue()
                                );
                                for (Row row : xssfSheet) {
                                    XSSFRow xssfRow = (XSSFRow) row;
                                    XSSFCell xssfCell = xssfRow.getCell(0);
                                    if (xssfCell == null || xssfCell.getCellType().equals(CellType.BLANK)) {
                                        Object value = Utils.getCellValue(xssfRow.getCell(1),
                                                                          inputWorkbook.getCreationHelper()
                                                                                       .createFormulaEvaluator()
                                        );
                                        if (value == null) {
                                            continue;
                                        }
                                        Object nameValue = Utils.getCellValue(xssfRow.getCell(2),
                                                                              inputWorkbook.getCreationHelper()
                                                                                           .createFormulaEvaluator()
                                        );
                                        String name;
                                        if (nameValue != null) {
                                            name = nameValue.toString();
                                        }
                                        else {
                                            name = "";
                                            logger.fatal("Unable to get item name");
                                            System.exit(1);
                                        }
                                        Object qtyValue = Utils.getCellValue(xssfRow.getCell(col),
                                                                             inputWorkbook.getCreationHelper()
                                                                                          .createFormulaEvaluator()
                                        );
                                        int qty;
                                        if (qtyValue instanceof Integer) {
                                            qty = (int) qtyValue;
                                        }
                                        else if (qtyValue != null) {
                                            try {
                                                qty = (int) Double.parseDouble(qtyValue.toString());
                                            }
                                            catch (NumberFormatException e) {
                                                qty = 0;
                                                logger.fatal("Unable to get Item Qty", e);
                                                System.exit(1);
                                            }
                                        }
                                        else {
                                            qty = 0;
                                            logger.fatal("Unable to get Item Qty");
                                            System.exit(1);
                                        }
                                        if (qty != 0) {
                                            itemGroup.addItem(name, qty);
                                        }
                                    }
                                    else {
                                        break;
                                    }
                                }
                                itemGroups.add(itemGroup);
                            }
                        }
                    }
                    XSSFSheet ns_item_group = workbook.createSheet("NS Item Group");
                    itemCount = 0;
                    for (ItemGroup itemGroup : itemGroups) {
                        if (itemGroup.getItems().size() > itemCount) {
                            itemCount = itemGroup.getItems().size();
                        }
                    }
                    headerRow = ns_item_group.createRow(0);
                    headerCell = headerRow.createCell(0);
                    headerCell.setCellValue("Item Group");
                    headerCell.setCellStyle(headerStyle);
                    headerCell = headerRow.createCell(1);
                    headerCell.setCellValue("Description");
                    headerCell.setCellStyle(headerStyle);
                    headerCell = headerRow.createCell(2);
                    headerCell.setCellValue("Member Item");
                    headerCell.setCellStyle(headerStyle);
                    headerCell = headerRow.createCell(3);
                    headerCell.setCellValue("Member Item Quantity");
                    headerCell.setCellStyle(headerStyle);
                    rowCount = 1;
                    for (ItemGroup itemGroup : itemGroups) {
                        Map<String, Integer> items = itemGroup.getItems();
                        for (String name : items.keySet()) {
                            XSSFRow row = ns_item_group.createRow(rowCount);
                            XSSFCell cell = row.createCell(0);
                            cell.setCellValue(itemGroup.getKey());
                            cell = row.createCell(1);
                            cell.setCellValue(itemGroup.getDescription());
                            cell = row.createCell(2);
                            cell.setCellValue(name);
                            cell = row.createCell(3);
                            cell.setCellValue(items.get(name));
                            rowCount++;
                        }
                    }
                    for (int i = 0; i < Utils.getLastColumn(ns_item_group); i++) {
                        ns_item_group.autoSizeColumn(i);
                    }
                    opcPackage.revert();
                    try {
                        String configuratorName = configurator.getName().split(" Configurator")[0].toLowerCase();
                        FileOutputStream fileOutputStream = new FileOutputStream(Paths.get("./" +
                                                                                           configuratorName +
                                                                                           "_bom_map.xlsx")
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
                else {
                    logger.fatal("Unable to find master sheet");
                    System.exit(1);
                }
            }
        }
    }

    private int getIntValue(Object o) {
        if (o instanceof Double) {
            double aDouble = (double) o;
            return (int) aDouble;
        }
        else if (o instanceof Integer) {
            return (int) o;
        }
        else {
            String str = o.toString();
            if (str != null) {
                str = str.strip();
                if (str.contains(".")) {
                    try {
                        double val = Double.parseDouble(str);
                        return (int) val;
                    }
                    catch (NumberFormatException e) {
                        logger.fatal("Invalid value", e);
                        System.exit(1);
                        return Integer.MIN_VALUE;
                    }
                }
                else {
                    try {
                        return Integer.parseInt(str);
                    }
                    catch (NumberFormatException e) {
                        logger.fatal("Invalid value", e);
                        System.exit(1);
                        return Integer.MIN_VALUE;
                    }
                }
            }
            else {
                logger.fatal("Invalid value");
                System.exit(1);
                return Integer.MIN_VALUE;
            }
        }
    }

}

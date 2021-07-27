package org.conversion.tasks.assemblies;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;

/**
 * @author Nicholas Curl
 */
public class AssemblyItem extends Item {

    /**
     * The instance of the logger
     */
    private static final Logger logger = LogManager.getLogger(AssemblyItem.class);

    private final String description;
    private final int    qty;

    public AssemblyItem(int itemNum, String key, String description, int qty) {
        super(itemNum,key);
        this.description = description;
        this.qty = qty;
    }

    public int getQty() {
        return qty;
    }

    public String getDescription() {
        return description;
    }

}

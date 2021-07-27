package org.conversion.tasks.assemblies;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;

/**
 * @author Nicholas Curl
 */
public class Item {

    /**
     * The instance of the logger
     */
    private static final Logger logger = LogManager.getLogger(Item.class);

    private final int itemNum;
    private final String itemKey;

    public Item (int itemNum, String itemKey){
        this.itemNum = itemNum;
        this.itemKey = itemKey;
    }

    public int getItemNum() {
        return itemNum;
    }

    public String getKey() {
        return itemKey;
    }
}

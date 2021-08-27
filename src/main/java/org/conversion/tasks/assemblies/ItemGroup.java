package org.conversion.tasks.assemblies;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;

import java.util.HashMap;
import java.util.Map;

/**
 * @author Nicholas Curl
 */
public class ItemGroup {

    /**
     * The instance of the logger
     */
    private static final Logger logger = LogManager.getLogger(ItemGroup.class);

    private final String               key;
    private final String               description;
    private final Map<String, Integer> items;

    public ItemGroup(String key, String description) {
        this.key = key;
        this.description = description;
        this.items = new HashMap<>();
    }

    public String getDescription() {
        return description;
    }

    public String getKey() {
        return key;
    }

    public Map<String, Integer> getItems() {
        return items;
    }

    @Override
    public String toString() {
        return "ItemGroup{" +
               "key='" + key + '\'' +
               '}';
    }

    public void addItem(String name, int qty) {
        this.items.put(name, qty);
    }
}

package org.conversion.tasks.assemblies;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;

import java.util.Objects;

/**
 * @author Nicholas Curl
 */
public class AssemblyMaster extends Item {

    /**
     * The instance of the logger
     */
    private static final Logger logger = LogManager.getLogger(AssemblyMaster.class);

    public AssemblyMaster(String key, int itemNum) {
        super(itemNum,key);
    }

    @Override
    public boolean equals(Object o) {
        if (this == o) {
            return true;
        }
        if (o == null || getClass() != o.getClass()) {
            return false;
        }
        AssemblyMaster that = (AssemblyMaster) o;
        return getItemNum() == that.getItemNum() && getKey().equalsIgnoreCase(that.getKey());
    }

    @Override
    public int hashCode() {
        return Objects.hash(getKey(), getItemNum());
    }
}

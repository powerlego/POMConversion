package org.conversion;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.conversion.tasks.Configurators;
import org.conversion.tasks.Packages;

/**
 * @author Nicholas Curl
 */
public class Main {

    /**
     * The instance of the logger
     */
    private static final Logger logger = LogManager.getLogger(Main.class);

    public static void main(String[] args) {
        new Packages().run();
        new Configurators().run();
    }

}

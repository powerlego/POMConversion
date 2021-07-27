package org.conversion.sql;

import com.microsoft.sqlserver.jdbc.SQLServerDriver;
import org.apache.commons.dbcp2.BasicDataSource;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;

import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.SQLException;

/**
 * @author Nicholas Curl
 */
public class DBConnectionPool {

    private static final BasicDataSource ds = new BasicDataSource();
    /**
     * The instance of the logger
     */
    private static final Logger logger = LogManager.getLogger(DBConnectionPool.class);

    static {
        try {
            DriverManager.registerDriver(new SQLServerDriver());
        }
        catch (SQLException e) {
            logger.fatal("Unable to register driver", e);
            System.exit(-1);
        }
        ds.setUrl("jdbc:sqlserver://192.168.20.232\\SQLEXPRESS:1433;database=POR3");
        ds.setUsername("dataprocessing");
        ds.setPassword("dataprocessing");
        ds.setMinIdle(5);
        ds.setMaxIdle(10);
        ds.setMaxOpenPreparedStatements(100);
    }

    public static Connection getConnection() throws SQLException{
        return ds.getConnection();
    }

    private DBConnectionPool(){}
}

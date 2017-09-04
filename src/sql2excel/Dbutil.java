package sql2excel;

import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.SQLException;

public class Dbutil {
	public static Connection getConnection() {
		Connection conn=null;
		
	    String JDBC_DRIVER="oracle.jdbc.driver.OracleDriver";
	    String DB_URL="jdbc:oracle:thin:@localhost:1521:XE";
	    String USER="info";
	    String PASS="123456";
	    
	    // 加载数据库连接驱动
	    try {
	        Class.forName(JDBC_DRIVER);
	    } catch (ClassNotFoundException e) {
	    // TODO Auto-generated catch block
	        e.printStackTrace();
	    }
	    // 获取数据库的连接对象
	    try {
	        conn = DriverManager.getConnection(DB_URL, USER, PASS);
	        System.out.println("数据库连接建立成功...");
	    } catch (SQLException e) {
	    // TODO Auto-generated catch block
	        e.printStackTrace();
	    }
	    // 返回连接对象
	    return conn;
	}
	public static void close(Connection c) {
		if (c != null) {
		try {
		c.close();
		} catch (Throwable e) {
		 
		e.printStackTrace();
		}
		}
		}
}

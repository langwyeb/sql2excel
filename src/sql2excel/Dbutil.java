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
	    
	    // �������ݿ���������
	    try {
	        Class.forName(JDBC_DRIVER);
	    } catch (ClassNotFoundException e) {
	    // TODO Auto-generated catch block
	        e.printStackTrace();
	    }
	    // ��ȡ���ݿ�����Ӷ���
	    try {
	        conn = DriverManager.getConnection(DB_URL, USER, PASS);
	        System.out.println("���ݿ����ӽ����ɹ�...");
	    } catch (SQLException e) {
	    // TODO Auto-generated catch block
	        e.printStackTrace();
	    }
	    // �������Ӷ���
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

package com.excel.util;

import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.SQLException;

/**
 * JDBC数据库连接工具类
 * @author Joker
 *
 */
public class JdbcUtil {

	private static String Url = "jdbc:oracle:thin:@172.23.20.153:1521:cottonx";
	private static Connection conn = null;
	
	/**
	 * 打开数据库连接
	 * @return
	 */
	public static Connection getConn(){
		try {
			Class.forName("oracle.jdbc.driver.OracleDriver");
			conn = DriverManager.getConnection(Url, "CenterUser", "Cen_Ac45$_SDHW0S_YUR0DF_#$wX3O");
			return conn;
		} catch (Exception e) {
			System.err.println("数据库连接打开失败");
			e.printStackTrace();
		}
		
		return null;
	}
	
	/**
	 * 关闭数据库连接
	 */
	public static void closeConn(){
		try {
			if(null != conn){
				conn.close();
			}
		} catch (SQLException e) {
			System.err.println("数据库连接关闭失败");
			e.printStackTrace();
		}
	}
}

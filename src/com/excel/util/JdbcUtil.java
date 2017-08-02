package com.excel.util;

import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.SQLException;

/**
 * JDBC���ݿ����ӹ�����
 * @author Joker
 *
 */
public class JdbcUtil {

	private static String Url = "jdbc:oracle:thin:@172.23.20.153:1521:cottonx";
	private static Connection conn = null;
	
	/**
	 * �����ݿ�����
	 * @return
	 */
	public static Connection getConn(){
		try {
			Class.forName("oracle.jdbc.driver.OracleDriver");
			conn = DriverManager.getConnection(Url, "CenterUser", "Cen_Ac45$_SDHW0S_YUR0DF_#$wX3O");
			return conn;
		} catch (Exception e) {
			System.err.println("���ݿ����Ӵ�ʧ��");
			e.printStackTrace();
		}
		
		return null;
	}
	
	/**
	 * �ر����ݿ�����
	 */
	public static void closeConn(){
		try {
			if(null != conn){
				conn.close();
			}
		} catch (SQLException e) {
			System.err.println("���ݿ����ӹر�ʧ��");
			e.printStackTrace();
		}
	}
}

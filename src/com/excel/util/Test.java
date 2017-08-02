package com.excel.util;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.Connection;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;

public class Test {

	private static final String SQL1 = "SELECT TABLE_NAME, TABLE_TYPE, COMMENTS FROM USER_TAB_COMMENTS";
	private static final String SQL2 = "select distinct a.column_name as COLUMN_NAME, a.data_type as dataType, a.DATA_LENGTH as dataLength, b.comments as COMMENTS from User_Tab_Cols a, USER_COL_COMMENTS b where b.column_name = a.column_name and b.table_name = a.table_name and a.table_name =";
	private static final String SQL3 = "SELECT t.table_name ����,"+
       "t.colUMN_NAME �ֶ�,"+
       "t1.COMMENTS ������,"+
       "case when t.COLUMN_NAME in("+
         "select col.column_name"+     
" from user_constraints con,  user_cons_columns col"+     
" where con.constraint_name = col.constraint_name"+     
" and con.constraint_type='P' and col.TABLE_NAME=t.table_name) then '����' else '' end as ����,"+
"t.data_default Ĭ��ֵ,"+
       "case when t.NULLABLE='N' then '����' when t.NULLABLE='Y' then '' end as �Ƿ�Ϊ��,"+
       "t.DATA_TYPE �ֶ�����,"+
"case when t.char_used is not null then t.char_length  WHEN t.DATA_TYPE IN ('CLOB','BLOB','TIMESTAMP(6)','DATE') THEN t.DATA_LENGTH else t.DATA_PRECISION end as ���� ,"+
"t.DATA_SCALE ����"+
  " FROM User_Tab_Cols t, User_Col_Comments t1"+
" WHERE t.table_name = t1.table_name"+
    " AND t.column_name = t1.column_name AND t.table_name=";
	
	public static void main(String[] args) {
		System.out.println("����ʼ");
		HSSFWorkbook wb = new HSSFWorkbook();
		ResultSet rs = null;
		try {
		    HSSFCellStyle linkStyle = wb.createCellStyle();
            /*����Font*/
            HSSFFont cellFont= wb.createFont();
            linkStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
            linkStyle.setFillForegroundColor(HSSFColor.YELLOW.index);
            linkStyle.setFont(cellFont);
            
            HSSFCellStyle headStyle = wb.createCellStyle();
            headStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
            headStyle.setFillForegroundColor(HSSFColor.PALE_BLUE.index);
		    
		    wb.createSheet("LOGLIST");
			Statement st = JdbcUtil.getConn().createStatement();
			rs = st.executeQuery(SQL1);
			
			HSSFSheet sheet = wb.createSheet("TAB");
			sheet.setColumnWidth(0, 5000);  
			sheet.setColumnWidth(1, 4000);
			sheet.setColumnWidth(2, 3500);  
			sheet.setColumnWidth(3, 3500);
			
			sheet.createRow(0);
			HSSFRow row = sheet.createRow(1);
			row.setRowStyle(headStyle);
			HSSFCell cell = row.createCell(0);
			HSSFCell cell1 = row.createCell(1);
			cell.setCellValue("����");
			cell1.setCellValue("ע��");
			cell.setCellStyle(headStyle);
			cell1.setCellStyle(headStyle);
			HSSFRow isol = sheet.createRow(2);
			HSSFCell isolCell = isol.createCell(0);
            isolCell.setCellValue("----");
            isolCell.setCellStyle(linkStyle);
            isol.setRowStyle(linkStyle);
			
			HSSFSheet sheet_col = wb.createSheet("COL");
			
			sheet_col.setColumnWidth(0, 5000);  
			sheet_col.setColumnWidth(1, 5000);
			sheet_col.setColumnWidth(2, 6000);  
			sheet_col.setColumnWidth(3, 3500);
			sheet_col.setColumnWidth(4, 3500);
			sheet_col.setColumnWidth(5, 3500);
			sheet_col.setColumnWidth(6, 3500);
			sheet_col.setColumnWidth(7, 3500);
			sheet_col.setColumnWidth(8, 3500);
			sheet_col.setColumnWidth(9, 3500);
			sheet_col.setColumnWidth(10, 3500);
			sheet_col.setColumnWidth(11, 3500);
			sheet_col.setColumnWidth(12, 3500);
			
			sheet_col.createRow(0);
			HSSFRow row_col = sheet_col.createRow(1);
			row_col.setRowStyle(headStyle);
			HSSFCell cell_col = row_col.createCell(0);
			HSSFCell cell_col1 = row_col.createCell(1);
			HSSFCell cell_col2 = row_col.createCell(2);
			HSSFCell cell_col3 = row_col.createCell(3);
			HSSFCell cell_col4 = row_col.createCell(4);
			HSSFCell cell_col5 = row_col.createCell(5);
			HSSFCell cell_col6 = row_col.createCell(6);
			HSSFCell cell_col7 = row_col.createCell(7);
			HSSFCell cell_col8 = row_col.createCell(8);
			HSSFCell cell_col9 = row_col.createCell(9);
			HSSFCell cell_col10 = row_col.createCell(10);
			HSSFCell cell_col11 = row_col.createCell(11);
			HSSFCell cell_col12 = row_col.createCell(12);
			
			cell_col.setCellValue("����");
			cell_col1.setCellValue("�ֶ�");
			cell_col2.setCellValue("�ֶ�������");
			cell_col3.setCellValue("����");
			cell_col4.setCellValue("Ĭ��ֵ");
			cell_col5.setCellValue("��Ϊ��");
			cell_col6.setCellValue("��������");
			cell_col7.setCellValue("���ݳ���");
			cell_col8.setCellValue("���ݾ���");
			cell_col9.setCellValue("�汾");
			cell_col10.setCellValue("˵��");
			cell_col11.setCellValue("��������");
			cell_col12.setCellValue("��չ����");
			
			cell_col.setCellStyle(headStyle);
			cell_col1.setCellStyle(headStyle);
			cell_col2.setCellStyle(headStyle);
			cell_col3.setCellStyle(headStyle);
			cell_col4.setCellStyle(headStyle);
			cell_col5.setCellStyle(headStyle);
			cell_col6.setCellStyle(headStyle);
			cell_col7.setCellStyle(headStyle);
			cell_col8.setCellStyle(headStyle);
			cell_col9.setCellStyle(headStyle);
			cell_col10.setCellStyle(headStyle);
			cell_col11.setCellStyle(headStyle);
			cell_col12.setCellStyle(headStyle);
			
			int i = 1;
			while(rs.next() ){
				String tableName = rs.getString("TABLE_NAME");
				String comments = rs.getString("COMMENTS");
				HSSFRow r = sheet.createRow(i+2);
				HSSFCell c = r.createCell(0);
				HSSFCell c1 = r.createCell(1);
				c1.setCellValue(comments);
				c.setCellValue(tableName);
				getList(tableName, sheet_col );
				
				HSSFRow row_isolate = sheet_col.createRow(row_num);
				HSSFCell cell_isolate = row_isolate.createCell(0);
				row_isolate.setRowStyle(linkStyle);
				cell_isolate.setCellValue("----");
				cell_isolate.setCellStyle(linkStyle);
				row_num++;
				i++;
			}
			wb.createSheet("INDEX");
			try {
				FileOutputStream os = new FileOutputStream("e:\\jianyanzhongxin.xls");  
				wb.write(os);  
				os.close();
			} catch (FileNotFoundException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}  
			
			System.out.println("ִ�н���");
			
			JdbcUtil.closeConn();
		} catch (SQLException e) {
			e.printStackTrace();
		}
	}
	
	private static int row_num=2;
	
	private static void getList(String tableName, HSSFSheet sheet) {
		ResultSet rs = null;

		try {
		    Connection conn= JdbcUtil.getConn();
			Statement st = conn.createStatement();
			rs = st.executeQuery(SQL3 + "'" + tableName + "'");
			
			while(rs.next()){
				String table_name = rs.getString("����");
				String table_column = rs.getString("�ֶ�");
				String column_name = rs.getString("������");
				String keys = rs.getString("����");
				String defualt = rs.getString("Ĭ��ֵ");
				String nullable = rs.getString("�Ƿ�Ϊ��");
				String column_type = rs.getString("�ֶ�����");
				String column_length = rs.getString("����");
				String column_scale = rs.getString("����");
				
				HSSFRow r = sheet.createRow(row_num);
				HSSFCell c = r.createCell(0);
				HSSFCell c1 = r.createCell(1);
				HSSFCell c2 = r.createCell(2);
				HSSFCell c3 = r.createCell(3);
				HSSFCell c4 = r.createCell(4);
				HSSFCell c5 = r.createCell(5);
				HSSFCell c6 = r.createCell(6);
				HSSFCell c7 = r.createCell(7);
				HSSFCell c8 = r.createCell(8);
				
				c.setCellValue(table_name);
				c1.setCellValue(table_column);
				c2.setCellValue(column_name);
				c3.setCellValue(keys);
				c4.setCellValue(defualt);
				c5.setCellValue(nullable);
				c6.setCellValue(column_type);
				c7.setCellValue(column_length);
				c8.setCellValue(column_scale);
				
				row_num++;
			}
			conn.close();
		} catch (SQLException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}

}

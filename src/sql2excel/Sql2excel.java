package sql2excel;

import java.io.File;
import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.ResultSet;

import jxl.Workbook;
import jxl.format.CellFormat;
import jxl.format.Colour;
import jxl.write.Label;
import jxl.write.WritableCellFormat;
import jxl.write.WritableFont;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;

public class Sql2excel {
	
	public static void main(String[] args) {
		//createXls();
		modifyXls();
	}

	//����1 �������ļ�
	public static void createXls() {
		try {
			WritableWorkbook book=Workbook.createWorkbook(new File("sql2excel.xls"));
			WritableSheet sheet = book.createSheet("sheet1", 0);
			
			//���ø�ʽ ������ɫ����
			WritableFont font1 = new WritableFont(WritableFont.createFont("����"), 10, WritableFont.BOLD);
			font1.setColour(Colour.BLACK);
			WritableCellFormat format1 = new WritableCellFormat(font1);
			format1.setAlignment(jxl.format.Alignment.CENTRE);
			format1.setVerticalAlignment(jxl.format.VerticalAlignment.CENTRE);
			
			Label labelA = new Label(0, 0, "USER_NAME",format1);
			Label labelB = new Label(1, 0, "USER_PWD",format1);
			Label labelC = new Label(2, 0, "TEL",format1);
			
			sheet.addCell(labelA);
			sheet.addCell(labelB);
			sheet.addCell(labelC);
			
			Connection conn=Dbutil.getConnection();
			PreparedStatement pstmt=conn.prepareStatement("select * from USERINFO");
			ResultSet result = pstmt.executeQuery();
			
			int i=0;
			while(result.next()) {
				Label labelAi = new Label(0, i+1, result.getString(1));
				Label labelBi = new Label(1, i+1, result.getString(2));
				Label labelCi = new Label(2, i+1, result.getString(3));
				sheet.addCell(labelAi);
				sheet.addCell(labelBi);
				sheet.addCell(labelCi);
				i++;
			}
			
			book.write();
			book.close();
			System.out.println("�����ļ��ɹ�!");
			pstmt.close();
			conn.close();
		} catch (Exception e) {
			System.out.println(e);
		}
	}
	
	//����2 ͨ���޸��ļ��������ļ�
	public static void modifyXls() {
		try {
			Workbook book_src=Workbook.getWorkbook(new File("sql2excel.xls"));
			//�������޸ĵĸ���
			WritableWorkbook book=Workbook.createWorkbook(new File("sql2excel_new.xls"),book_src);
			WritableSheet sheet=book.getSheet(0);
			
			Connection conn=Dbutil.getConnection();
			PreparedStatement pstmt=conn.prepareStatement("select * from USERINFO");
			ResultSet result = pstmt.executeQuery();
			
			int i=0;
			while(result.next()) {
				//��ȡ��Ԫ���ʽ
				CellFormat cfa=sheet.getCell(0,i+1).getCellFormat();
				CellFormat cfb=sheet.getCell(1,i+1).getCellFormat();
				CellFormat cfc=sheet.getCell(2,i+1).getCellFormat();
				//д����
				Label labelAi = new Label(0, i+1, result.getString(1),cfa);
				Label labelBi = new Label(1, i+1, result.getString(2),cfb);
				Label labelCi = new Label(2, i+1, result.getString(3),cfc);
				sheet.addCell(labelAi);
				sheet.addCell(labelBi);
				sheet.addCell(labelCi);
				i++;
			}
			
			book.write();
			book.close();
			System.out.println("�޸��ļ��ɹ�!");
			pstmt.close();
			conn.close();
		} catch (Exception e) {
			System.out.println(e);
		}
	}
}

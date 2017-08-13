package com.info.TestBase;

import java.io.FileInputStream;
import java.io.IOException;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
//import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class DEMO {

	public static void main(String[] args) throws Exception {

		try {

			Class forName = Class.forName("com.mysql.jdbc.Driver");
			Connection con = null;
			con = DriverManager.getConnection("jdbc:mysql://localhost/testdb", "root", "System@1234");
			con.setAutoCommit(false);
			PreparedStatement pstm = null;
			FileInputStream fis = new FileInputStream("E:\\Seleium_data\\Dummy_Data_Framwork_Using_Selenium\\TestDataFile\\Datasheet.xlsx");
			XSSFWorkbook workbook = new XSSFWorkbook(fis);
			XSSFSheet spreadsheet = workbook.getSheetAt(0);
			
			Sheet sheet = workbook.getSheetAt(0);
			Row row;
			for (int i = 1; i <= sheet.getLastRowNum(); i++) {
				row = (Row) sheet.getRow(i);
				String name = row.getCell(0).getStringCellValue();
				System.out.println(name+"\t\t");
				String add = row.getCell(1).getStringCellValue();
				System.out.println(add+"\t\t");
				String city = row.getCell(2).getStringCellValue();
				System.out.println(city+"\t\t");
				String mobile = row.getCell(3).getStringCellValue();
				System.out.println(mobile+"\t\t");
				String sql = "INSERT INTO test (name, address, city, mobile) VALUES('" + name + "','" + add + "'," + city + ",'" + mobile + "')";
				pstm = (PreparedStatement) con.prepareStatement(sql);
				pstm.execute();
				System.out.println("Import rows " + i);
			}
			con.commit();
			pstm.close();
			con.close();
			fis.close();
			System.out.println("Success import excel to mysql table");
		} catch (IOException e) {
		}
	}
}
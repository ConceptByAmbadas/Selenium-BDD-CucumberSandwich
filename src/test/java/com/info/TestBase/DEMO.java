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

	Connection con=null;
	PreparedStatement pst=null;
	
	public String name;
	public String city;
	public String add;
	public String mobile;
	
	public void readWriteData()
	{
		try
		{
		Class forName = Class.forName("com.mysql.jdbc.Driver");
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
			 name = row.getCell(0).getStringCellValue();
			System.out.println(name+"\t\t");
			 add = row.getCell(1).getStringCellValue();
			System.out.println(add+"\t\t");
			city = row.getCell(2).getStringCellValue();
			System.out.println(city+"\t\t");
			 mobile = row.getCell(3).getStringCellValue();
			System.out.println(mobile+"\t\t");
			pstm = con.prepareStatement("INSERT INTO test (name, address, city, mobile) VALUES(?,?,?,?)");
			//pstm.setString(1, id);
			pstm.setString(1, name);
			pstm.setString(2, add);
			pstm.setString(3, city);
			pstm.setString(4, mobile);
			int n=pstm.executeUpdate();
			System.out.println("Import rows " + n);
		}
		con.commit();
		pstm.close();
		con.close();
		fis.close();
		System.out.println("Success import excel to mysql table");
		}
		catch(Exception Ex)
		{System.out.println("Exception is"+Ex.getMessage());}
	}
	public static void main(String[] args) throws Exception {

		
		DEMO d1=new DEMO();
		d1.readWriteData();
		
		
	}
}
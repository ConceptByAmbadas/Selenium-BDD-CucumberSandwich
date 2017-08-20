package com.info.TestBase;

import java.io.FileInputStream;
import java.io.IOException;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.Time;
import java.sql.Timestamp;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Date;

import org.apache.poi.poifs.filesystem.POIFSFileSystem;
//import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class DEMO {

	public Connection con=null;
	public PreparedStatement pstm=null;

	public String name;
	public String city;
	public String add;
	public int mobile;
	public int sr_no;
	public String time_taken;



	public DEMO()
	{
		try
		{
			Class forName = Class.forName("com.mysql.jdbc.Driver");
			con = DriverManager.getConnection("jdbc:mysql://localhost/testdb", "root", "System@1234");
			con.setAutoCommit(false);
		}
		catch(Exception es)
		{
			System.out.println("Exception is"+es.getMessage());
		}
	}
	public void readWriteData()
	{
		try
		{
			FileInputStream fis = new FileInputStream("E:\\Seleium_data\\Dummy_Data_Framwork_Using_Selenium\\TestDataFile\\Datasheet.xlsx");
			XSSFWorkbook workbook = new XSSFWorkbook(fis);
			
			Sheet sheet = workbook.getSheetAt(0);
			Row row;
			for (int i = 1; i <= sheet.getLastRowNum(); i++) {
				row = (Row) sheet.getRow(i);

				sr_no = (int)row.getCell(0).getNumericCellValue();
				System.out.println(sr_no+"\t\t");
				name = row.getCell(1).getStringCellValue();
				System.out.println(name+"\t\t");
				add = row.getCell(2).getStringCellValue();
				System.out.println(add+"\t\t");
				//city = row.getCell(3).getStringCellValue();
				//System.out.println("City is"+city+"\t\t");


				city = row.getCell(3).getStringCellValue();
				Timestamp ts = Timestamp.valueOf(city);
				System.out.println(ts);

				mobile = (int)row.getCell(4).getNumericCellValue();
				System.out.println(mobile+"\t\t");

				time_taken= row.getCell(5).getStringCellValue();
				Time.valueOf(time_taken);
				DateFormat sdf = new SimpleDateFormat("hh:mm:ss");
				Date date = sdf.parse(time_taken);
				System.out.println("Value is"+Time.valueOf(time_taken));


				pstm = con.prepareStatement("INSERT INTO test (name, address, city, mobile,time_taken) VALUES(?,?,?,?,?)");
				//pstm.setInt(1, id);
				pstm.setString(1, name);
				pstm.setString(2, add);
				//pstm.setString(3,  city);
				pstm.setTimestamp(3, ts);
				pstm.setInt(4, mobile);
				///pstm.setString(5, );
				pstm.setTime(5, Time.valueOf(time_taken));
				int n=pstm.executeUpdate();
				System.out.println("Data Import successfully" + n);

			}
			con.commit();
			pstm.close();
			con.close();
			fis.close();

			//System.out.println("Success import excel to mysql table");
		}
		catch(Exception Ex)
		{Ex.printStackTrace();}
	}
	public static void main(String[] args) throws Exception {


		DEMO d1=new DEMO();
		d1.readWriteData();


	}
}
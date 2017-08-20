package com.info.TestBase;
import java.io.File;
import java.io.FileInputStream;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.util.Iterator;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
public class EXCEL {

	static XSSFRow row;
	public  String name;
	public  String add;
	public  String city;
	public  String mobile;

	public  String id=(String.format("%02d", 0));

	public void readWriteData()
	{
		try
		{
			Class forName = Class.forName("com.mysql.jdbc.Driver");
			Connection con = null;
			con = DriverManager.getConnection("jdbc:mysql://localhost/testdb", "root", "System@1234");
			//con.setAutoCommit(false);
			PreparedStatement pstm = null;



			String excelFilePath = "E:\\Seleium_data\\Dummy_Data_Framwork_Using_Selenium\\TestDataFile\\Datasheet.xlsx";
			FileInputStream inputStream = new FileInputStream(new File(excelFilePath));

			Workbook workbook = new XSSFWorkbook(inputStream);
			Sheet firstSheet = workbook.getSheetAt(0);
			Iterator<Row> iterator = firstSheet.iterator();

			while (iterator.hasNext()) {
				Row nextRow = iterator.next();
				Iterator<Cell> cellIterator = nextRow.cellIterator();

				while (cellIterator.hasNext()) {
					Cell cell = cellIterator.next();

					if(cell.getColumnIndex()==0)//for example of c
					{
						name=cell.getStringCellValue();
						System.out.print(name);
					}

					if(cell.getColumnIndex()==1)//for example of c
					{
						add=cell.getStringCellValue();
						System.out.print(add);
					}

					if(cell.getColumnIndex()==2)//for example of c
					{
						city=cell.getStringCellValue();
						System.out.print(city);
					}

					if(cell.getColumnIndex()==3)//for example of c
					{
						mobile=cell.getStringCellValue();
						System.out.print(mobile);
					}

				     
					System.out.print(" - ");
				}
				System.out.println();
			}

			// String sql = "INSERT INTO test (id,name, address, city, mobile) VALUES(?,?,?,?,?)";
			pstm = con.prepareStatement("INSERT INTO test (name, address, city, mobile) VALUES(?,?,?,?)");
			//pstm.setString(1, id);
			pstm.setString(1, name);
			pstm.setString(2, add);
			pstm.setString(3, city);
			pstm.setString(4, mobile);

			System.out.println("mobile is"+mobile);
			int i=pstm.executeUpdate();
			System.out.println("Import rows " + i);

			workbook.close();
			inputStream.close();
			con.close();
			pstm.close();



		}
		catch(Exception ex)
		{System.out.println("Exception is"+ex.getMessage());}

	}
	public static void main(String[] args) throws Exception 
	{


		EXCEL Ex=new EXCEL();

		Ex.readWriteData();

	}
}

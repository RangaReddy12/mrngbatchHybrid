package utilities;

import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class ExcelFileUtil {
Workbook wb;
//constructor for reading excel path
public ExcelFileUtil(String Excelpath)throws Throwable
{
	FileInputStream fi = new FileInputStream(Excelpath);
	wb = WorkbookFactory.create(fi);
}
//get row count from sheet
public int rowCount(String sheetname)
{
	return wb.getSheet(sheetname).getLastRowNum();
}
//method for reading cell data
public String getCellData(String sheetName,int row,int column)
{
	String data="";
	if(wb.getSheet(sheetName).getRow(row).getCell(column).getCellType()==Cell.CELL_TYPE_NUMERIC)
	{
		//read integert type cell and store
		int celldata = (int) wb.getSheet(sheetName).getRow(row).getCell(column).getNumericCellValue();
		//convert integer type to string
		data =String.valueOf(celldata);
	}
	else
	{
		data =wb.getSheet(sheetName).getRow(row).getCell(column).getStringCellValue();
	}
	return data;
	
}
public void setCellData(String sheetName,int row,int column,String status,String WriteExcel)throws Throwable
{
	//get sheet from row
	Sheet ws = wb.getSheet(sheetName);
	//get row from sheet
	Row rowNum =ws.getRow(row);
	//create cell
	Cell cell = rowNum.createCell(column);
	//write status
	cell.setCellValue(status);
	if(status.equalsIgnoreCase("Pass"))
	{
		CellStyle style = wb.createCellStyle();
		Font font = wb.createFont();
		font.setColor(IndexedColors.GREEN.getIndex());
		font.setBold(true);
		font.setBoldweight(Font.BOLDWEIGHT_BOLD);
		style.setFont(font);
		rowNum.getCell(column).setCellStyle(style);
	}
	else if(status.equalsIgnoreCase("Fail"))
	{
		CellStyle style = wb.createCellStyle();
		Font font = wb.createFont();
		font.setColor(IndexedColors.RED.getIndex());
		font.setBold(true);
		font.setBoldweight(Font.BOLDWEIGHT_BOLD);
		style.setFont(font);
		rowNum.getCell(column).setCellStyle(style);
	}
	else if(status.equalsIgnoreCase("Blocked"))
	{
		CellStyle style = wb.createCellStyle();
		Font font = wb.createFont();
		font.setColor(IndexedColors.BLUE.getIndex());
		font.setBold(true);
		font.setBoldweight(Font.BOLDWEIGHT_BOLD);
		style.setFont(font);
		rowNum.getCell(column).setCellStyle(style);
	}
	FileOutputStream fo = new FileOutputStream(WriteExcel);
	wb.write(fo);
			
}
public static void main(String[] args)throws Throwable {
	ExcelFileUtil xl = new ExcelFileUtil("D:/Sample.xlsx");
	int rc =xl.rowCount("EmpData");
	System.out.println("No of rows are::"+rc);
	for(int i=1;i<=rc;i++)
	{
		String fname = xl.getCellData("EmpData", i, 0);
		String mname = xl.getCellData("EmpData", i, 1);
		String lname = xl.getCellData("EmpData", i, 2);
		String eid = xl.getCellData("EmpData", i, 3);
		System.out.println(fname+"    "+mname+"    "+lname+"    "+eid);
		//xl.setCellData("EmpData", i, 4, "Pass", "D:/Results.xlsx");
		//xl.setCellData("EmpData", i, 4, "Fail", "D:/Results.xlsx");
		xl.setCellData("EmpData", i, 4, "Blocked", "D:/Results.xlsx");
	}
	
}
}













package utilities;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class ExcelFileUtil {
	
Workbook wb;
//constructor for reading excel path
public ExcelFileUtil(String ExcelPath) throws Throwable
{
	FileInputStream fi=new FileInputStream(ExcelPath);
	wb=WorkbookFactory.create(fi);
}
//methods for counting rows
public int rowCount(String SheetName)
{
	return
	wb.getSheet(SheetName).getLastRowNum();
}
//read cell data
public String getCellData(String SheetName,int Row,int Column)
{
	String data="";
	if(wb.getSheet(SheetName).getRow(Row).getCell(Column).getCellType()==Cell.CELL_TYPE_NUMERIC)
	{
		int celldata=(int) wb.getSheet(SheetName).getRow(Row).getCell(Column).getNumericCellValue();
		data=String.valueOf(celldata);
}
	else
	{
		data=wb.getSheet(SheetName).getRow(Row).getCell(Column).getStringCellValue();
	}
	return data;

}
//method for set cell data
public void setCellData(String SheetName,int Row,int Column,String Status,String WriteExcel) throws Throwable
{
	//get sheet from wb
	Sheet ws=wb.getSheet(SheetName);
	//get row from sheet
	org.apache.poi.ss.usermodel.Row rowNum=ws.getRow(Row);
	//create cell
	Cell cell=rowNum.createCell(Column);
	//write status
	cell.setCellValue(Status);
	if(Status.equalsIgnoreCase("Pass"))
	{
		CellStyle style=wb.createCellStyle();
		Font font=wb.createFont();
		font.setColor(IndexedColors.GREEN.getIndex());
		font.setBold(true);
		font.setBoldweight(Font.BOLDWEIGHT_BOLD);
		style.setFont(font);
		rowNum.getCell(Column).setCellStyle(style);
		
	}
	else if(Status.equalsIgnoreCase("Fail"))
	{
		CellStyle style=wb.createCellStyle();
		Font font=wb.createFont();
		font.setColor(IndexedColors.RED.getIndex());
		font.setBold(true);
		font.setBoldweight(Font.BOLDWEIGHT_BOLD);
		style.setFont(font);
		rowNum.getCell(Column).setCellStyle(style);
		}
	else if(Status.equalsIgnoreCase("Blocked"))
		{
			CellStyle style=wb.createCellStyle();
			Font font=wb.createFont();
			font.setColor(IndexedColors.BLUE.getIndex());
			font.setBold(true);
			font.setBoldweight(Font.BOLDWEIGHT_BOLD);
			style.setFont(font);
			rowNum.getCell(Column).setCellStyle(style);
		}
	FileOutputStream fo=new FileOutputStream(WriteExcel);
	wb.write(fo);
	
}
}

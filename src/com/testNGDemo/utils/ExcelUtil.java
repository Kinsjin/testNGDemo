package com.testNGDemo.utils;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelUtil {
	//Sheet
	//WorkBook
	//Cell
	//Row
	private static XSSFSheet excelWSheet;
	private static XSSFWorkbook excelWorkBook;
	private static XSSFCell cell;
	private static XSSFSheet row;
	boolean isExcel2007;
	
	public static void main(String args[]){
		System.out.println("HelloWorld");
		String path="C:/Users/jinwx/Desktop/123.xlsx";
		ExcelUtil util= new ExcelUtil();
		boolean fileStatus=util.excelExist(path);
		if(!fileStatus){
			System.out.println("it is not excel");
		}else
			util.setExcelFile(path,"Sheet1");
			//util.dataReader(7, 1);
			System.out.println(util.dataReader(0, 1));
			
	}
	//check the file exist;
	public boolean excelExist(String path){
		//boolean fileExist=false;
		File file= new File(path);
		if(file.exists()&&file!=null){
		//get the prefix
			String prefix=path.substring(path.lastIndexOf("."));
			if(path==null||prefix.equals(".xlsx")||prefix.equals(".XLSX")){
				//check the filename end with xlsx
				isExcel2007=true;
			}
		}else isExcel2007=false;
		return isExcel2007;
	}
	//identified Excel file
	public void setExcelFile(String path,String sheetName){
		FileInputStream excelFile;
		try {
			//instance the excel file;
			excelFile=new FileInputStream(path);
			//instance the excel file's workbook
			excelWorkBook=new XSSFWorkbook(excelFile);
			//instace the sheet;
			excelWSheet= excelWorkBook.getSheet(sheetName);
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}catch(IOException e){
			// TODO Auto-generated catch block
						e.printStackTrace();
		}
	}
	//Data Reader
	public String dataReader(int rownum,int cellnum){
		try{
			//set the cell num;
			cell= excelWSheet.getRow(rownum).getCell(cellnum);
			//jugement the valuetype of cell;
			String cellData=cell.getCellType()==XSSFCell.CELL_TYPE_STRING?cell.getStringCellValue()
					:String.valueOf(Math.round(cell.getNumericCellValue()));
			System.out.println(cellData);
			return cellData;
			
		}catch(Exception e){
			return "";
		}
	}
}

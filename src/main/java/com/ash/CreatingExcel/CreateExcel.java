package com.ash.CreatingExcel;

import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class CreateExcel {
	final static String FILE_NAME="New.xlsx";
	
	public static void main(String[] args) {
		try(XSSFWorkbook xssfWorkbook = new XSSFWorkbook();) {
		XSSFSheet createSheet = xssfWorkbook.createSheet("Sheet");
		Object [][] file= {
				{"Name","LastName","MiddleName","Age"},
				{"Ashique","Ali","Mahar",12},
				{"Siraj","haq","Chandio",13}
		};
		int rowNum=0;
		for(Object [] rows:file) {
			XSSFRow createRow = createSheet.createRow(rowNum++);
			int columnIndex=0;
			for (Object row : rows) {
				XSSFCell createCell = createRow.createCell(columnIndex++);
				if (row instanceof String) {
					createCell.setCellValue((String)row);
				}else if (row instanceof Integer) {
					createCell.setCellValue((Integer)row);
				}
				
			}
		}
		System.out.println("Excel Created");
		FileOutputStream fStream=new FileOutputStream(FILE_NAME);
		xssfWorkbook.write(fStream);
		fStream.close();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		//createSheet.clos
	}
}

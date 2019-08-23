package org.cts.test.MavenProject;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class Bank {

	public static void main(String[] args) throws IOException {
		//Set the file location
		File loc=new File("C:\\Users\\raja sekar\\eclipse-workspace\\MavenDay1\\MavenProject\\Driver\\Pradeep.xlsx");
		FileInputStream stream=new FileInputStream(loc);
		
		//Workbook
		Workbook w=new XSSFWorkbook(stream);
		//Sheet
		Sheet s=w.getSheet("Sheet1");
		//Row
	for (int i = 0; i < s.getPhysicalNumberOfRows(); i++) {
		Row r = s.getRow(i);
		for (int j = 0; j < r.getPhysicalNumberOfCells(); j++) {
			Cell c = r.getCell(j);
		
		int type=c.getCellType();
		
		//type==1//String format
		if(type==1)
			
		{
			
			String stringCells=c.getStringCellValue();
			System.out.println(stringCells);
		}
		//type==0//date format
		else if(type==0)
		{
			if(DateUtil.isCellDateFormatted(c))
			{
				Date dateCell=c.getDateCellValue();
				SimpleDateFormat sim=new SimpleDateFormat("dd/mm/yyyy");
				String f=sim.format(dateCell);
				System.out.println(f);
			}
			else
			{
				double numericCells=c.getNumericCellValue();
				//typecast
			
				long l=(long)numericCells;
				//to convert long into String
				String v=String.valueOf(l);
				System.out.println(v);
			}
			
			
			}
	}
	
	}
	}
}

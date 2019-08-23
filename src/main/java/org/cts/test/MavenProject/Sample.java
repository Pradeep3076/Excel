package org.cts.test.MavenProject;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Sample {

	public static void main(String[] args) throws IOException {
		//Set the file location
		File loc=new File("C:\\Users\\raja sekar\\eclipse-workspace\\MavenDay1\\MavenProject\\Driver\\Karthik.xlsx");
		FileInputStream stream=new FileInputStream(loc);
		
		//Workbook
		Workbook w=new XSSFWorkbook(stream);
		
		Sheet s=w.getSheet("Sheet1");
		
		Row r=s.getRow(1);
		
		Cell c=r.getCell(0);
		
		String s1 = c.getStringCellValue();
		if(s1.equals("loc")) {
			c.setCellValue("Farith");
		}
		FileOutputStream o=new FileOutputStream(loc);
		w.write(o);
		System.out.println("Done");
		}

}

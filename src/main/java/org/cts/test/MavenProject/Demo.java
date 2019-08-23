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

public class Demo {

	public static void main(String[] args) throws IOException {
		//Set the file loc
		File loc=new File("C:\\Users\\raja sekar\\eclipse-workspace\\MavenDay1\\MavenProject\\testdata\\SSD.xlsx");
		FileInputStream stream=new FileInputStream(loc);
		
		//Workbook
		Workbook w=new XSSFWorkbook(stream);
		//Sheet
		Sheet s=w.createSheet("Set1");
		//Row
		Row r=s.createRow(5);
		//Cell
		Cell c=r.createCell(7);
		c.setCellValue("Sairam");
		
		FileOutputStream n=new FileOutputStream(loc);
		w.write(n);
		System.out.println("Completed");
		}
	}


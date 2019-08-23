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

public class Driver {

	public static void main(String[] args) throws IOException {
		//Set the file location
		File m=new File("C:\\Users\\raja sekar\\eclipse-workspace\\MavenDay1\\MavenProject\\testdata\\Arun.xlsx");
		FileInputStream stream=new FileInputStream(m);
		
		//Workbook
		Workbook w=new XSSFWorkbook(stream);
		//Sheet
		Sheet s=w.createSheet("List1");
		//Row
		Row r=s.createRow(3);
		//Cell
		Cell c=r.createCell(3);
		c.setCellValue("Ganesh");
		
		FileOutputStream o=new FileOutputStream(m);
		w.write(o);
		System.out.println("Success");
		}

}

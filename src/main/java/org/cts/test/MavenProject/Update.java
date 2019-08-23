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

public class Update {

	public static void main(String[] args) throws IOException {
		//Set the file loc
		File loc=new File("C:\\Users\\raja sekar\\eclipse-workspace\\MavenDay1\\MavenProject\\testdata\\Arun.xlsx");
		FileInputStream stream=new FileInputStream(loc);
		
		//Workbook
		Workbook w=new XSSFWorkbook(stream);
		//Sheet
		Sheet s=w.getSheet("List1");
		//Row
		Row r=s.getRow(3);
		//Cell
		Cell c=r.getCell(3);
		String s1=c.getStringCellValue();
		if(s1.equals("Ganesh"))
		{
			c.setCellValue("Kannan");
		}
		FileOutputStream o=new FileOutputStream(loc);
		w.write(o);
		System.out.println("Done");
		}

}

package org.Fb;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

import javax.swing.text.Element;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.Select;

import io.github.bonigarcia.wdm.WebDriverManager;

public class FaceBookMv {
public static void main(String[] args) throws IOException, InvalidFormatException {
	WebDriverManager.chromedriver().setup();
	WebDriver driver=new ChromeDriver();
	driver.get("http:demo.automationtesting.in/Register.html");

	Workbook Workbook=new XSSFWorkbook();
	
	WebElement skills = driver.findElement(By.id("Skills"));
	Select select=new Select(skills);
	List<WebElement> options = select.getOptions();
	
	for(int i=0; i<options.size();i++) {
		WebElement element = options.get(i);
	    String text = element.getText();
	    System.out.println(text);
}
	int count = options.size();
    System.out.println("Count Size:"+count);
    
    
    Sheet createSheet = Workbook.createSheet();
    
    Row createRow = createSheet.createRow(0);
    Cell createCell = createRow.createCell(0);
    createCell.setCellValue("Adobe InDesign");
    

    Row createRow1 = createSheet.createRow(1);
    Cell createCell2 = createRow1.createCell(0);
    createCell2.setCellValue("Adobe Photoshop");
    
    Row createRow3 = createSheet.createRow(2);
    Cell createCell3 = createRow3.createCell(0);
    createCell3.setCellValue("Analytics");
    
    Row createRow4 = createSheet.createRow(3);
    Cell createCell4 = createRow4.createCell(0);
    createCell4.setCellValue("Android");
    

    Row createRow5 = createSheet.createRow(4);
    Cell createCell5 = createRow5.createCell(0);
    createCell5.setCellValue("APIs");
    
    Row createRow6 = createSheet.createRow(5);
    Cell createCell6 = createRow6.createCell(0);
    createCell6.setCellValue("Art Design");
    
    Row createRow7 = createSheet.createRow(6);
    Cell createCell7 = createRow7.createCell(0);
    createCell7.setCellValue("AutoCAD");
    
    Row createRow8 = createSheet.createRow(7);
    Cell createCell8 = createRow8.createCell(0);
    createCell8.setCellValue("Backup Management");
    
    Row createRow9 = createSheet.createRow(8);
    Cell createCell9 = createRow9.createCell(0);
    createCell9.setCellValue("C");
    

    Row createRow10 = createSheet.createRow(9);
    Cell createCell10 = createRow10.createCell(0);
    createCell10.setCellValue("C++");
    
    Row createRow11 = createSheet.createRow(10);
    Cell createCell11 = createRow11.createCell(0);
    createCell11.setCellValue("Certifications");
    
    Row createRow12 = createSheet.createRow(11);
    Cell createCell12 = createRow12.createCell(0);
    createCell12.setCellValue("Client Server");
    
    Row createRow13= createSheet.createRow(12);
    Cell createCell13 = createRow13.createCell(0);
    createCell13.setCellValue("Client Support");
    
    Row createRow14 = createSheet.createRow(13);
    Cell createCell14 = createRow14.createCell(0);
    createCell14.setCellValue("Configuration");
    
    Row createRow15 = createSheet.createRow(14);
    Cell createCell15 = createRow15.createCell(0);
    createCell15.setCellValue("Content Managment");
    

    Row createRow16 = createSheet.createRow(15);
    Cell createCell16 = createRow16.createCell(0);
    createCell16.setCellValue("Content Management Systems (CMS)");
    
    Row createRow17 = createSheet.createRow(16);
    Cell createCell17 = createRow17.createCell(0);
    createCell17.setCellValue("Corel Draw");
     
    File file=new File("C:\\Users\\vinot\\eclipse-workspace\\FacebookMaven\\src\\test\\java\\Excel\\kani.xcel.xlsx");
    FileOutputStream out=new FileOutputStream(file);
    Workbook.write(out);
    
    driver.close();
    
    
    
    
    
    
    
    
    
    
    
    
    
}}


package com.xls;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
 
public class ConvertXLSXToTXT {
 
    private static void convertSelectedSheetInXLXSFileToCSV(File xlsxFile, File outputFile) throws Exception {
    	int sheetIdx=0;
        FileInputStream fileInStream = new FileInputStream(xlsxFile);
        //File outputFile=null;
        FileOutputStream fos = new FileOutputStream(outputFile);
 
        // Open the xlsx and get the requested sheet from the workbook
        XSSFWorkbook wb = new XSSFWorkbook(fileInStream);
        XSSFSheet sh = wb.getSheetAt(sheetIdx);
        int starRow = sh.getFirstRowNum();
        int endRow = sh.getLastRowNum();
        
    
        // Iterate through all the rows in the selected sheet
        //Iterator<Row> rowIterator = selSheet.iterator();
       // List[] ag = new List[];
        ArrayList<String> ag1=new ArrayList<String>();
        ArrayList<String> attributes=new ArrayList<String>();
        ArrayList<String> valueset=new ArrayList<String>();
        
        ArrayList<String> finalset=new ArrayList<String>();
       
        for(int i = starRow + 1; i <=endRow; i++) {
        	  String ag = wb.getSheetAt(0).getRow(i).getCell(0).toString();
        	 // System.out.println("ag:" +ag.getStringCellValue());
        	  ag1.add(ag);
        	 
        	 String att = wb.getSheetAt(0).getRow(i).getCell(1).toString();
        	 attributes.add(att);
        	  
        	 String vs = wb.getSheetAt(0).getRow(i).getCell(2).toString();
        	 valueset.add(vs);
        	 
        	 finalset.addAll(ag1);
        	 finalset.addAll(attributes);
        	 finalset.addAll(valueset);
        	 
        	// System.out.println("Value is ::"+wb.getSheetAt(0).getRow(i+1).getCell(0).toString());
        	/*if(ag== wb.getSheetAt(0).getRow(i+1).getCell(0).toString())
        	 {
        		 continue;
        	 }
        	else
        	{
        		
        	}*/
            }
        
        System.out.println("finalse" +finalset);
           // fos.write(sb.toString().getBytes());
        	//fos.close();
        
        wb.close();
       
}
	
 
    public static void main(String[] args) throws Exception {
        File myFile = new File("C:\\InputFiles\\input.xlsx");
       // int sheetIdx = 0; // 0 for first sheet
 
        //convertSelectedSheetInXLXSFileToCSV(myFile, sheetIdx);
        File outputFile = new File("C:\\InputFiles\\output.txt");
        convertSelectedSheetInXLXSFileToCSV(myFile, outputFile);
    }
}

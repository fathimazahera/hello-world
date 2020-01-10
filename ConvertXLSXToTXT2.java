package com.xls;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.*;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
 
public class ConvertXLSXToTXT2 {
 
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
        
        //Map<String, List<Integer>> map = new HashMap<String, List<Integer>>();
        Map<String, List<String>> map = new HashMap<String, List<String>>();
      // ArrayList<String> finalset=new ArrayList<String>();
       
        for(int i = starRow + 1; i <=endRow; i++) {
        	  String ag = wb.getSheetAt(0).getRow(i).getCell(0).toString();
        	 // System.out.println("ag:" +ag.getStringCellValue());
        	 // ag1.add(ag);
        	 
        	
        	 if(!(map.containsKey(ag)))
        	 {
        		 ArrayList<String> finalset=new ArrayList<String>();
        		 String att = wb.getSheetAt(0).getRow(i).getCell(1).toString();
	        	 String vs = wb.getSheetAt(0).getRow(i).getCell(2).toString();
            	
            	 finalset.add(att);
            	 finalset.add(vs);
            	 
        		 map.put(ag, finalset);
        	 }
        	 else
        	 {
        		 
        		 List<String> AttnVs = map.get(ag);
        		 String att = wb.getSheetAt(0).getRow(i).getCell(1).toString();
	        	 String vs = wb.getSheetAt(0).getRow(i).getCell(2).toString();
	        	 AttnVs.add(att);
	        	 AttnVs.add(vs);
        		 map.put(ag, AttnVs);
        		 
        	 }
        	 
        	// System.out.println("map key::" +ag + "map values::" +map.get(ag));
            }
         System.out.println("map::" +map);
       
      //  System.out.println("finalse" +finalset);
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

package com.xls;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.ObjectOutputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.*;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
 
public class ConvertXLSXToTXT3 {
 
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
       
        for(int i = starRow; i <=endRow; i++) {
        	  String ag = wb.getSheetAt(0).getRow(i).getCell(0).toString();
        	 // System.out.println("ag:" +ag.getStringCellValue());
        	 // ag1.add(ag);
        	 
        	
        	 if(!(map.containsKey(ag)))
        	 {
        		 ArrayList<String> finalset=new ArrayList<String>();
        		 String att = wb.getSheetAt(0).getRow(i).getCell(1).toString();
	        	 String vs = wb.getSheetAt(0).getRow(i).getCell(2).toString();
	        	 String att3 = wb.getSheetAt(0).getRow(i).getCell(3).toString();
	        	 String att4 = wb.getSheetAt(0).getRow(i).getCell(4).toString();
	        	 String att5 = wb.getSheetAt(0).getRow(i).getCell(5).toString();
	        	 String att6= wb.getSheetAt(0).getRow(i).getCell(6).toString();
	        	 String att7= wb.getSheetAt(0).getRow(i).getCell(7).toString();
	        	 String att8= wb.getSheetAt(0).getRow(i).getCell(8).toString();
	        	 String att9= wb.getSheetAt(0).getRow(i).getCell(9).toString();
	        	 String att10= wb.getSheetAt(0).getRow(i).getCell(10).toString();
	        	 String att11= wb.getSheetAt(0).getRow(i).getCell(11).toString();
	        	 String att12= wb.getSheetAt(0).getRow(i).getCell(12).toString();
            	 finalset.add(att);
            	 finalset.add(vs);
            	 finalset.add(att3);
            	 finalset.add(att4);
            	 finalset.add(att5);
            	 finalset.add(att6);
            	 finalset.add(att7);
            	 finalset.add(att8);
            	 finalset.add(att9);
            	 finalset.add(att10);
            	 finalset.add(att11);
            	 finalset.add(att12);
            	 
            	 
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
        	 
        	
            }
         System.out.println("map::" +map);
         System.out.println("length:" +map.size());
         
         StringBuffer sb = new StringBuffer();
         for(int i=0; i<map.keySet().size(); i++)
         {
        	
        	 File file = new File("C:\\InputFiles\\"+map.keySet().toArray()[i]+".txt");
        	 FileOutputStream f = new FileOutputStream(file);
        	 Object key = map.keySet().toArray()[i];
        	 System.out.println("keyset:" +key+ "values:" +map.get(key));
           	 sb.append(key);
        	 sb.append(map.get(key));
        	 f.write(sb.toString().getBytes());
        	 f.close(); 
        	 
         }
             wb.close();
       
}
	
 
    public static void main(String[] args) throws Exception {
        File myFile = new File("C:\\InputFiles\\AttributeGroupFile.xlsx");
       // int sheetIdx = 0; // 0 for first sheet
 
        //convertSelectedSheetInXLXSFileToCSV(myFile, sheetIdx);
        File outputFile = new File("C:\\InputFiles\\output.txt");
        convertSelectedSheetInXLXSFileToCSV(myFile, outputFile);
    }
}

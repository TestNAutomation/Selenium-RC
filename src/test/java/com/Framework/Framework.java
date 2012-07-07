package com.Framework;

//package com.example.tests;

import com.thoughtworks.selenium.*;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.junit.After;
import org.junit.Before;
import org.junit.Test;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.regex.Pattern;

public class Framework extends SeleneseTestCase {
	// Declare our variables (GLOBAL)
	public String vLoan, vTerm, vRate, vMonth, vYear, vPtax, vPMI;
	public String vURL1, vURL2;
	public int xRows, xCols;
	public String xData[][]; 
	@Before
	public void setUp() throws Exception {
		String xPath = "C:/Selenium/Jul5/mc-data.xls";
		xlRead(xPath);
		myprint("Rows are " + xRows);
		myprint("Cols are " + xCols);
		selenium = new DefaultSelenium("localhost", 1236, "*chrome", "http://www.mortgagecalculator.org/");
		selenium.start();
	}

	@Test
	public void testMc5() throws Exception {
		String vApp1, vApp2, vResult; // Output variables
		vURL1 = "http://www.mortgagecalculator.org/";
		vURL2 = "http://www.mortgagecalculatorplus.com/";
		for (int i = 1; i < xRows; i++ ){
			//if (xData[i][10] == "Y") {
			if (xData[i][10].equals("Y")) {
				// Load values into the parameters from the excel (xData)
				myprint("Reading data for Row " + i);
				vLoan = xData[i][0];
				vTerm = xData[i][1];
				vRate = xData[i][2];
				vMonth = xData[i][3];
				vYear = xData[i][4];
				vPtax = xData[i][5];
				vPMI = xData[i][6];
		
				vApp1 = app1();
				vApp2 = app2();
				//if (vApp1 == vApp2){
				if (vApp1.equals(vApp2)){
					vResult = "Pass";
					System.out.println("PASS");
				}else {
					vResult = "Fail";
					System.out.println("FAIL");
				}
				xData[i][7] = vApp1;
				xData[i][8] = vApp2;
				xData[i][9] = vResult;
			}
		}
	}

	@After
	public void tearDown() throws Exception {
		// write results back to the excel
		String xRPath = "C:/Selenium/Jul5/mc-data-res.xls";
		xlwrite(xRPath, xData);
		selenium.stop();
	}
	
	public String app1(){
		myprint("Inside the App1 Fun");
		selenium.open(vURL1);
		selenium.type("param[principal]", vLoan);
		selenium.type("param[term]", vTerm);
		selenium.type("param[interest_rate]", vRate);
		selenium.select("param[start_month]", vMonth);
		selenium.select("param[start_year]", vYear);
		selenium.type("param[property_tax]", vPtax);
		selenium.type("param[pmi]", vPMI);
		selenium.click("css=input[type=submit]");
		selenium.waitForPageToLoad("25000");
		String mc1 = selenium.getText("css=td > h3");
		System.out.println("value from app1 is " + mc1);
		return mc1;
	}
	
	public String app2(){
		myprint("Inside the App2 Fun");
		selenium.open(vURL2);
		selenium.type("param[principal]", vLoan);
		selenium.type("param[term]", vTerm);
		selenium.type("param[interest_rate]", vRate);
		selenium.select("param[start_month]", vMonth);
		selenium.select("param[start_year]", vYear);
		selenium.type("param[property_tax]", vPtax);
		selenium.type("param[pmi]", vPMI);
		selenium.click("css=input[type=submit]");
		selenium.waitForPageToLoad("25000");
		String mc2 = selenium.getText("css=td > h3");
		System.out.println("value from app2 is " + mc2);
		return mc2;
	}
	
	public void xlRead(String sPath) throws Exception{
		File myxl = new File(sPath);
		FileInputStream myStream = new FileInputStream(myxl);
		
		HSSFWorkbook myWB = new HSSFWorkbook(myStream);
		//HSSFSheet mySheet = new HSSFSheet(myWB);
		//HSSFSheet mySheet = myWB.getSheetAt(0);	// Referring to 1st sheet
		HSSFSheet mySheet = myWB.getSheetAt(2);	// Referring to 3rd sheet
		//int xRows = mySheet.getLastRowNum()+1;
		//int xCols = mySheet.getRow(0).getLastCellNum();
		xRows = mySheet.getLastRowNum()+1;
		xCols = mySheet.getRow(0).getLastCellNum();
		myprint("Rows are " + xRows);
		myprint("Cols are " + xCols);
		//String[][] xData = new String[xRows][xCols];
		xData = new String[xRows][xCols];
      for (int i = 0; i < xRows; i++) {
	           HSSFRow row = mySheet.getRow(i);
	            for (int j = 0; j < xCols; j++) {
	               HSSFCell cell = row.getCell(j); // To read value from each col in each row
	               String value = cellToString(cell);
	               xData[i][j] = value;
	               System.out.print(value);
	               System.out.print("@@");
	               }
	            System.out.println("");
	            
	        }	
		
	}
	
	public static String cellToString(HSSFCell cell) {
	// This function will convert an object of type excel cell to a string value
      int type = cell.getCellType();
      Object result;
      switch (type) {
          case HSSFCell.CELL_TYPE_NUMERIC: //0
              result = cell.getNumericCellValue();
              break;
          case HSSFCell.CELL_TYPE_STRING: //1
              result = cell.getStringCellValue();
              break;
          case HSSFCell.CELL_TYPE_FORMULA: //2
              throw new RuntimeException("We can't evaluate formulas in Java");
          case HSSFCell.CELL_TYPE_BLANK: //3
              result = "-";
              break;
          case HSSFCell.CELL_TYPE_BOOLEAN: //4
              result = cell.getBooleanCellValue();
              break;
          case HSSFCell.CELL_TYPE_ERROR: //5
              throw new RuntimeException ("This cell has an error");
          default:
              throw new RuntimeException("We don't support this cell type: " + type);
      }
      return result.toString();
  }
	public void xlwrite(String xlPath, String[][] xldata) throws Exception {
		System.out.println("Inside XL Write");
  	File outFile = new File(xlPath);
      HSSFWorkbook wb = new HSSFWorkbook();
         // Make a worksheet in the XL document created
      /*HSSFSheet osheet = wb.setSheetName(1,"TEST");*/
      HSSFSheet osheet = wb.createSheet("TESTRESULTS");
      // Create row at index zero ( Top Row)
  	for (int myrow = 0; myrow < xRows; myrow++) {
  		//System.out.println("Inside XL Write");
	        HSSFRow row = osheet.createRow(myrow);
	        // Create a cell at index zero ( Top Left)
	        for (int mycol = 0; mycol < xCols; mycol++) {
	        	HSSFCell cell = row.createCell(mycol);
	        	// Lets make the cell a string type
	        	cell.setCellType(HSSFCell.CELL_TYPE_STRING);
	        	// Type some content
	        	cell.setCellValue(xldata[myrow][mycol]);
	        	//System.out.print("..." + xldata[myrow][mycol]);
	        }
	        //System.out.println("..................");
	        // The Output file is where the xls will be created
	        FileOutputStream fOut = new FileOutputStream(outFile);
	        // Write the XL sheet
	        wb.write(fOut);
	        fOut.flush();
//		    // Done Deal..
	        fOut.close();
  	}
  }
	public void myprint(String mymessage){
		System.out.println(mymessage);
		System.out.println("~~~~~~~~~~~~~");
	}
}

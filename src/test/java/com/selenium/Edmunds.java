package com.selenium;

//Comment1 : 
//Instructs where the pre-defined methods are present.
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
//import java.util.regex.Pattern;

@SuppressWarnings("deprecation")
public class Edmunds extends SeleneseTestCase {
	public double fTotalDownApp, fTotalDownExp; 
	public int xRows, xCols;
	public String xData[][];
	@Before
	public void setUp() throws Exception {
//		System.out.println("Start setup");
		selenium = new DefaultSelenium("localhost", 1234, "*chrome", "http://www.edmunds.com/");
		//selenium = new DefaultSelenium("localhost", 1234, "*iehta", "http://www.edmunds.com/");
		selenium.start();
		xlRead("c:/Selenium/S29/edmunds_com.xls");
	}

	@Test
	public void testEd_tc1() throws Exception {
		// Paramterization - Initializing
		// Hard coding
		String vAfford, vZip, vRate, vTerm, vDown, vTradeIn;
		String vResult;
		String vOwed = "1200";
		int i=1;
		
		for(i=1; i<xRows; i=i+1){
			if (xData[i][10].equals("Y")) {
				vAfford = xData[i][0];
				vZip = xData[i][1];
				vRate = xData[i][3];
				vTerm = xData[i][2];
				vDown = xData[i][5];
				vTradeIn = xData[i][4];
				myPrintText("Iteration number is " + i);
				sTC1(vAfford);
				sTC2(vZip, vTerm, vRate, vTradeIn, vOwed, vDown);
				sTC3(vDown, vTradeIn, vOwed);
				myPrintText("Total Down needed is " + fTotalDownApp);
				myPrintText("Total Down needed Expected is " + fTotalDownExp);
				vResult = cmpDouble(fTotalDownApp,fTotalDownExp);
				myPrintText("Test Result is " + vResult);
				xData[i][6] = Double.toString(fTotalDownApp);
				xData[i][9] = vResult;
			} else {
				myPrintText("Skipped row number " + i );
			}
		}
	}
	

	@After
	public void tearDown() throws Exception {
		// Write results back to the Excel file
		xlwrite("c:/Selenium/S29/edmunds_com_results.xls", xData);
		// System.out.println("End setup");
		selenium.stop();
	}
	// My Custom Methods
	public void myPrintLine(){
		System.out.println("~~~~~~~~~~~~~~~~~");
	}

	public void myPrintText(String fPrint){
		System.out.println(fPrint);
	}
	
	public String cmpStrings(String fS1, String fS2){
		if(fS1.equals(fS2)){
			return "Pass";
		} else {
			return "Fail";
		}
	}
	public void sTC1(String fAfford){
	//	0. Step 0 
		selenium.open("/calculators/");
	//	1. Enter Affordability - Where, What, Value
		selenium.type("//*[@id='calc_input1']", fAfford);
	//	2. Click on go button
		selenium.click("//button[@name='Go']");
		selenium.waitForPageToLoad("20000");
	}
	public void sTC2(String fZip, String fTerm, String fRate, String fTradeIn, String fOwed, String fDown) throws InterruptedException{
		//		3.1 Check the main heading in the page
		
		System.out.println("Header text is " + selenium.getText("//html/body/div[4]/div[2]/div[2]/div/div/h1/span"));
//		3. Type the zip code.
		selenium.type("//*[@id='ac_zip_code']", fZip);
//		4. Check for the target monthly payment matches with step 1. - PENDING
		
//		5. Select a Loan Term
		selenium.select("//select[@name='ac_loan_term']", fTerm);
//		6. Finance Rate
		selenium.type("//*[@id='ac_market_finance_rate']", fRate);
//		7.1 Enter the trade-in
		selenium.type("//*[@id='ac_vehicle2_price']", fTradeIn);
//		7.2 Enter the trade-in amount we owe.
		selenium.type("//*[@id='ac_vehicle2_value_owed']", fOwed);
//		7.3 Down payment 
		selenium.type("//*[@id='ac_cash_down_payment']", fDown);
//		8. Click Calculate
		selenium.click("//*[@id='calculate-button']");
//		8.1 Wait for about 5 seconds before next step.
		Thread.sleep(7000);
	}
	public void sTC3(String fDown, String fTradeIn, String fOwed){
//		9.1 Capture the Total Down Payment
		String vTemp = selenium.getText("//*[@id='ac_total_down_payment_result']");
		vTemp = currencyToString(vTemp);
		fTotalDownApp = Double.parseDouble(vTemp);
		
//		9.2 Capture TMV price.
		myPrintText("TMV Value is " + selenium.getText("//*[@id='ac_max_tmv_result']"));
		
		//10.1 Compare Down
		fTotalDownExp = Double.parseDouble(fDown) + Double.parseDouble(fTradeIn) - Double.parseDouble(fOwed);
		
		//return fTotalDownApp;
		//return fTotalDownExp;
	}
	public void cmpStrings_old(String fS1, String fS2){
		if(fS1.equals(fS2)){
			myPrintText("Pass");
		} else {
			myPrintText("Fail");
		}
	}
	
	public String cmpDouble(double fS1, double fS2){
		if(fS1 == fS2){
			return "Pass";
		} else {
			return "Fail";
		}
	}
	
	public void xlRead(String sPath) throws Exception{
		File myxl = new File(sPath);
		FileInputStream myStream = new FileInputStream(myxl);
		
		HSSFWorkbook myWB = new HSSFWorkbook(myStream);
		HSSFSheet mySheet = myWB.getSheetAt(2);	// Referring to 3rd sheet
		xRows = mySheet.getLastRowNum()+1;
		xCols = mySheet.getRow(0).getLastCellNum();
		myPrintText("Rows are " + xRows);
		myPrintText("Cols are " + xCols);
		xData = new String[xRows][xCols];
     for (int i = 0; i < xRows; i++) {
	           HSSFRow row = mySheet.getRow(i);
	            for (int j = 0; j < xCols; j++) {
	               HSSFCell cell = row.getCell(j); // To read value from each col in each row
	               String value = cellToString(cell);
	               xData[i][j] = value;
	               System.out.print(value);
	               System.out.print("    ");
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
     HSSFSheet osheet = wb.createSheet("TESTRESULTS");
 	for (int myrow = 0; myrow < xRows; myrow++) {
	        HSSFRow row = osheet.createRow(myrow);
	        for (int mycol = 0; mycol < xCols; mycol++) {
	        	HSSFCell cell = row.createCell(mycol);
	        	cell.setCellType(HSSFCell.CELL_TYPE_STRING);
	        	cell.setCellValue(xldata[myrow][mycol]);
	        }
	        FileOutputStream fOut = new FileOutputStream(outFile);
	        wb.write(fOut);
	        fOut.flush();
	        fOut.close();
 	}
 }
	public String currencyToString(String vStr){
		String[] vStrTemp;
//		3.1 Check the main heading in the page
		vStrTemp = vStr.split("");
		String vOut="";
		for(int i =2; i < vStrTemp.length ; i++){
			vOut = vOut + vStrTemp[i];
		}
		String[] vOutTemp = vOut.split(",");
		String vOutNew = vOutTemp[0] + vOutTemp[1]; 
		return vOutNew;
	}
}

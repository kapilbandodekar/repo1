package com.nics.qa.util;

import java.io.*;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.Reporter;
import org.testng.annotations.DataProvider;

import com.nics.qa.base.TestBase;

public class ExcelUtils extends TestBase
{
	public static String[][] readExcelData(String sheetName, String filePath,
		String tableName) {
		String[][] testData = null;
		Workbook workBook = null;
		try 
		{
			InputStream inputStream = getInputFileStream(filePath);
			if (filePath.toLowerCase().endsWith("xlsx") == true) 
			{
				workBook = new XSSFWorkbook(inputStream);
			} 
			else if (filePath.toLowerCase().endsWith("xls") == true) 
			{
				workBook = new HSSFWorkbook(inputStream);
			}
			inputStream.close();
			Sheet sheet = workBook.getSheet(sheetName);
			Cell[] boundaryCells = findCell(sheet, tableName);
			Cell startCell = boundaryCells[0];
			Cell endCell = boundaryCells[1];
			int startRow = startCell.getRowIndex() + 1;
			int endRow = endCell.getRowIndex() - 1;
			int startCol = startCell.getColumnIndex() + 1;
			int endCol = endCell.getColumnIndex() - 1;
			System.out.println("Excel Array created with row size: "
					+ (endRow - startRow + 1) + " :: Column size: "
					+ (endCol - startCol + 1));
			testData = new String[endRow - startRow + 1][endCol - startCol + 1];
			for (int i = startRow; i < endRow + 1; i++) 
			{
				for (int j = startCol; j < endCol + 1; j++) 
				{
					if (sheet.getRow(i).getCell(j) == null) 
					{
						testData[i - startRow][j - startCol] = "";
					} 
					else 
					{
						testData[i - startRow][j - startCol] = sheet.getRow(i)
								.getCell(j).getStringCellValue();
					}
				}
			}
		} 
		catch (FileNotFoundException e) 
		{
			System.out.println("Could not read the Excel sheet");
			e.printStackTrace();
		}
		catch (IOException e) 
		{
			System.out.println("Could not read the Excel sheet");
			e.printStackTrace();
		}
		catch (Exception e) 
		{
			System.out.println("Exception in read the Excel sheet:"
					+ e.getMessage());
			e.printStackTrace();
		}
		return testData;
	}

	private static InputStream getInputFileStream(String fileName)
			throws FileNotFoundException 
	{
		ClassLoader loader = ExcelUtils.class.getClassLoader();
		InputStream inputStream = loader.getResourceAsStream(fileName);
		if (inputStream == null) {
			inputStream = new FileInputStream(new File(fileName));
		}
		return inputStream;
	}

	public static Cell[] findCell(Sheet sheet, String text) 
	{
		String pos = "start";
		Cell[] cells = new Cell[2];
		for (Row row : sheet) {
			for (Cell cell : row) {
				cell.setCellType(Cell.CELL_TYPE_STRING);
				if (text.equals(cell.getStringCellValue())) {
					if (pos.equalsIgnoreCase("start")) {
						cells[0] = (Cell) cell;
						pos = "end";
					} else {
						cells[1] = (Cell) cell;
					}
				}

			}
		}
		return cells;
	}

	public void updateAlfrescoArticleID(String sheetName, String filePath,
			String previousArticleID, String currentArticleID) 
	{
		try 
		{
			InputStream inp = new FileInputStream(filePath);
			Workbook wb = WorkbookFactory.create(inp);
			Sheet sheet = wb.getSheet(sheetName);
			for (Row row : sheet) 
			{
				for (Cell cell : row) 
				{
					cell.setCellType(Cell.CELL_TYPE_STRING);
					if (previousArticleID.equals(cell.getStringCellValue())) 
					{
						cell.setCellValue(currentArticleID);
					}
				}
			}
			// Write the output to a file
			FileOutputStream fileOut = new FileOutputStream(filePath);
			wb.write(fileOut);
			fileOut.close();
		} 
		catch (FileNotFoundException e) 
		{
			Reporter.log("Could not find the Excel sheet");
			e.printStackTrace();
		} 
		catch (IOException e) 
		{
			Reporter.log("Could not read the Excel sheet");
			e.printStackTrace();
		} 
		catch (Exception e) 
		{
			Reporter.log("Exception occured in updateAlfrescoArticleID: "
					+ e.getMessage());
			e.printStackTrace();
		}
	}
	
	public void write2DArrayToExcel(String[][] array, String sheetName,String filePath) 
	{
		InputStream inp = null;
		Workbook wb = null;
		Sheet sheet = null;
		try 
		{
			inp = new FileInputStream(System.getProperty("user.dir") + filePath);
			wb = WorkbookFactory.create(inp);
		} 
		catch (FileNotFoundException e) 
		{
			wb = new HSSFWorkbook();
		} 
		catch (Exception e) 
		{
			Reporter.log("Could not find the Excel sheet");
			e.printStackTrace();
		}
		try 
		{
			sheet = wb.createSheet(sheetName);
			for (int i = 0; i < array.length; i++) 
			{
				Row row = sheet.createRow(i);
				for (int j = 0; j < array[0].length; j++) 
				{
					Cell cell = row.createCell(j);
					cell.setCellType(Cell.CELL_TYPE_STRING);
					cell.setCellValue(array[i][j]);
				}
			}
			// Write the output to a file
			FileOutputStream fileOut = new FileOutputStream(System.getProperty("user.dir") + filePath);
			wb.write(fileOut);
			fileOut.close();
		} 
		catch (IOException e) 
		{
			Reporter.log("Could not read the Excel sheet");
			e.printStackTrace();
		} 
		catch (Exception e) 
		{
			Reporter.log("Exception occured in write2DArrayToExcel: "+ e.getMessage());
			e.printStackTrace();
		}
	}

//Data Provider
   @DataProvider(name="CreateOTW")
	public static Object[][] getTestData()
	{
		Object[][] testData=readExcelData("otwdetails",ExcelPaths.otwdetails,"OtwData");  
		return testData;
	}
   
   @DataProvider(name="CreateAejOTW")
	public static Object[][] getAejTestData()
	{
		Object[][] testData=readExcelData("otwdetails",ExcelPaths.otwdetails,"OtwAejData");  
		return testData;
	}
	
   @DataProvider(name="CreateEmeaOTW")
  	public static Object[][] getEmeaTestData()
  	{
  		Object[][] testData=readExcelData("otwdetails",ExcelPaths.otwdetails,"OtwEmeaData");  
  		return testData;
  	}
   
	@DataProvider(name="BusiApproval")
	public static Object[][] getBusiTestData()
	{
		Object[][] testData=readExcelData("busidata",ExcelPaths.busidata,"BusiData");  
		return testData;
	}
	
	@DataProvider(name="BusiAejApproval")
	public static Object[][] getBusiAejTestData()
	{
		Object[][] testData=readExcelData("busidata",ExcelPaths.busidata,"BusiAejData");  
		return testData;
	}
	
	@DataProvider(name="BusiEmeaApproval")
	public static Object[][] getBusiEmeaTestData()
	{
		Object[][] testData=readExcelData("busidata",ExcelPaths.busidata,"BusiAejData");  
		return testData;
	}
	
	@DataProvider(name="CompApproval")
	public static Object[][] getCompTestData()
	{
		Object[][] testData=readExcelData("compdata",ExcelPaths.compdata,"CompData");  
		return testData;
	}	
	
	@DataProvider(name="CompAejApproval")
	public static Object[][] getCompAejTestData()
	{
		Object[][] testData=readExcelData("compdata",ExcelPaths.compdata,"CompAejData");  
		return testData;
	}	
	
	@DataProvider(name="CompEmeaApproval")
	public static Object[][] getCompEmeaTestData()
	{
		Object[][] testData=readExcelData("compdata",ExcelPaths.compdata,"CompAejData");  
		return testData;
	}
	@DataProvider(name="CreateMSL")
	public static Object[][] getMslTestData()
	{
		Object[][] testData=readExcelData("mslDetails",ExcelPaths.mslDetails,"MSLData");  
		return testData;
	}
	@DataProvider(name="CreateAeJMSL")
	public static Object[][] getMslAeJTestData()
	{
		Object[][] testData=readExcelData("mslDetails",ExcelPaths.mslDetails,"MSLAeJData");  
		return testData;
	}
	@DataProvider(name="CreateEmeaMSL")
	public static Object[][] getMslEmeaTestData()
	{
		Object[][] testData=readExcelData("mslDetails",ExcelPaths.mslDetails,"MSLEmeaData");  
		return testData;
	}
	
	/*@DataProvider(name="SGMasEAPO")
	public static Object[][] addMasEaPo()
	{
		Object[][] testData=readExcelData("MasEaPoData",ExcelPaths.addMasEaPo,"SGMasEAPO");  
		return testData;
	}*/
}
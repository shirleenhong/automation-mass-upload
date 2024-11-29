package ServiceNow.CommonUtilities;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelFunctions 
{
	public static XSSFWorkbook workbook;
	public static XSSFSheet sheet;
	public static XSSFRow row;
	public static XSSFCell cell;
	
	public static String[] readData(String filePath, String sheetName) throws Exception{
		FileInputStream fs = new FileInputStream(filePath);
		workbook = new XSSFWorkbook(fs);
		Sheet sheet = workbook.getSheet(sheetName);
		Iterator<Row> iterator = sheet.iterator();
		String[] data= new String[sheet.getLastRowNum()+1];
		int counter=0;
		while(iterator.hasNext())
		{
			Row currentRow = iterator.next();
			Iterator<Cell> cellIterator = currentRow.iterator();
			data[counter]="";
			while (cellIterator.hasNext()) {
				Cell currentCell = cellIterator.next();
				if (currentCell.getCellTypeEnum() == CellType.STRING) 
				{
						data[counter]= data[counter]+currentCell.getStringCellValue() + "--";
				} 
				else if (currentCell.getCellTypeEnum() == CellType.NUMERIC) 
				{
					if(DateUtil.isCellDateFormatted(currentCell))
					{
						String d = currentCell.toString();
						data[counter]= data[counter]+d+ "--";
					}
					else
					{
					data[counter]= data[counter]+currentCell.getNumericCellValue() + "--";
				}
				}
				else if (currentCell.getCellTypeEnum() == CellType.FORMULA) 
				{
					FormulaEvaluator formulaEval = workbook.getCreationHelper().createFormulaEvaluator();
					String value=formulaEval.evaluate(currentCell).formatAsString();
					data[counter]= data[counter]+value + "--";
				}
				else if (currentCell.getCellTypeEnum() == CellType.BLANK)
				{
					currentCell.setCellValue("BLANK");
					data[counter]= data[counter]+currentCell.getStringCellValue() + "--";
				}	
			}
			counter++;//Moving to the next row
		}
		return data;
	}
	
	public static void writeDataAtCell(String fileName, String sheetName, int rowNum, int colNum, String data) throws Exception{
		//File file= new File(fileName);
		
		FileInputStream fis= new FileInputStream(fileName);
		workbook=new XSSFWorkbook(fis);
		sheet= workbook.getSheet(sheetName);
		
		row=sheet.getRow(rowNum);
		cell= row.createCell(colNum);
		cell.setCellValue(data);
		fis.close();
		FileOutputStream fos= new FileOutputStream(fileName);
		workbook.write(fos);
		fos.close();
	}
	
	public static String getData(int rwNum,int clNum){

		row  = sheet.getRow(rwNum);
		cell = row.getCell(clNum);
		
		String stringCellValue;
		double intCallValue;
		String data;
		
		try
	    {
			if (cell == null)
			{
				return "";
				
			}
			else
			{
				
				cell.setCellType(CellType.STRING);
				data = cell.getStringCellValue();
		        // parseInt(cell.getStringCellValue());
				stringCellValue = cell.getStringCellValue();
				data = stringCellValue;
			}
	        //return true;
	        
	    } catch (NumberFormatException ex)
	    {
	    	intCallValue = cell.getNumericCellValue();
	        data =  Double.toString(intCallValue);
	    }
		return data;
	}
	
	public  static void setPath(String path, String sheetName) throws IOException
	{

		FileInputStream fis = new FileInputStream(path);
		workbook = new XSSFWorkbook(fis);
		sheet = workbook.getSheet(sheetName);
	}
}


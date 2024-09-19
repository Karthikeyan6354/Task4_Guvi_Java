package task5.Task4_Guvi_Java;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import snippet.ReadingExcelSheet;
import snippet.WritingOperation;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

public class ReadWriteDateIntoExcel 
{         
        private static final String File_Name = "Utils//Sample.xlsx"; // Mentioning the file name, where going to write and read
	    
        public static void main(String[] args) throws IOException 
	    {
	        writeExcel();     //Calling the Write Method
	        readExcel();      //Calling the read method, once wrote the data's into the above excel
	    }

	    private static void writeExcel() throws IOException 
	    {
	        Workbook workbook = new XSSFWorkbook();   //Create a new workbook
	        Sheet sheet = workbook.createSheet("Data");   //Create new Work sheet in the workbook 
	        Object[][] data = { {"  Name", "       Age", " Email"},
	                            {"John Doe",   30, "John@test.com"},            //Setting the data's to be filled in the excel
	                            {"Jane Doe", 28, "John@test.com"},
	                            {"Bob Smith", 35, "Jacky@example.com"},
	                            {"Swapnil Mark",   37, "Swapnil@example.com"},};

	        int rowNum = 0;
	        for (Object[] rowData : data) 
	          {
	            Row row = sheet.createRow(rowNum++);   //Create a new row in the sheet
	            int colNum = 0;
	            for (Object field : rowData) 
	              {
	                Cell cell = row.createCell(colNum++);   //Create a new cell in the sheet 
	                 if (field instanceof String)          //Checking field object is an instance of String
	                      cell.setCellValue((String) field);
	                   else if (field instanceof Integer) 
	                       cell.setCellValue((Integer) field);
	               }
	           }
            try
            {           	
            FileOutputStream outputStream = new FileOutputStream(File_Name); // Create an output stream to write the workbook data to a excel file
	        workbook.write(outputStream);   //Write the data into the file
	        outputStream.close();   //closing the output stream
	        workbook.close();       //closing the workbook
	        System.out.println("Written the data's in excel");
            }
            catch (IOException e)  //Catch the exception
            {
            	e.getMessage();   //this will give the name of exception
            	e.printStackTrace();  //This will give information of the exception and error in page number
            }
      
	    }
	    
	   private static void readExcel() throws IOException 
	   {
	        try
	        {
		    FileInputStream inputStream = new FileInputStream(File_Name);  //Create Input stream to read the excel file
	        Workbook workbook = new XSSFWorkbook(inputStream);  // Create a Workbook instance for the file

	        Sheet sheet = workbook.getSheetAt(0);  //Getting first sheet of the workbook
	        for (Row row : sheet) 
	         {
	            for (Cell cell : row) 
	              {
	                switch (cell.getCellType())    //Process the cells based on the type
	                {
	                    case STRING:
	                        System.out.print(cell.getStringCellValue() + "\t");  //based on the data type, it will print the value with tab space
	                        break;
	                    case NUMERIC:
	                        System.out.print(cell.getNumericCellValue() + "\t");
	                        break;
	                    case BOOLEAN:
	                        System.out.print(cell.getBooleanCellValue() + "\t");
	                        break;
	                    default:
	                        System.out.print("\t");
	                }
	              }
	           System.out.println();   
	         }

	        inputStream.close();   //closing the input stream
	        workbook.close();   //closing the workbook
	        }
	       catch (IOException e)
	       {
	    	e.getMessage();
        	e.printStackTrace();
	      }
	   }
}

	        
	        


	


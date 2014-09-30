

import java.util.Arrays;

	import java.io.File;
	import java.io.FileInputStream;
	import java.io.FileNotFoundException;
	import java.io.IOException;

	import org.apache.poi.hssf.usermodel.HSSFCell;
	import org.apache.poi.hssf.usermodel.HSSFSheet;
	import org.apache.poi.hssf.usermodel.HSSFWorkbook;

	// Microsoft Xl sheet 97-2003 , .xls
	
	public class ReadDataSample {
		public static void main(String[] args) throws IOException 
		{
			String recData[][] = null;
			String secondData[][] = null;
			recData = readSheet("/Users/Fatema/Desktop/SampleData_1.xls", "Sheet1" );
			secondData = readSheet("/Users/Fatema/Desktop/SampleData_1.xls", "Sheet3" );
			int iRowCount = recData.length;
			int iColumnCount = recData[0].length;
			
			

			for(int i =0 ; i < iRowCount ; i++)
			{
				for(int j = 0; j < iColumnCount; j++)
				{
					
					
					System.out.println(recData[i][j]);
					System.out.println(secondData[i][j]);
					if(recData[i][j].equals(secondData[i][j]))
					{
						
						System.out.println(recData[i][j]+" equal "+secondData[i][j]);
					}
					else
					{
						System.out.println("not equal");
					}
				}
				
				
			}
			//System.out.println("sheet1 equal sheet2");	
		}

		/*	Method Name: readSheet
		 *  Method Description: Read data from Data Sheet
		 *  Arguments: dataTablePath = Path of the Data Sheet; sheetName = Name of Sheet
		 *  Created by: BSS 
		 *  Created Date: May 16 2014
		 *  Last Modified: May 16 2014 
		 */

		private static char[] getStringcellValue(int i) {
			// TODO Auto-generated method stub
			return null;
		}

		public static String[][] readSheet(String dataTablePath, String sheetName) throws IOException
		{
			String xlData[][] = null;

			/*Step 1: Access Data Table Path */
			File xlFile = new File(dataTablePath);

			/*step2: Access the Xl File*/
			FileInputStream xlDoc = new FileInputStream(xlFile);

			/*Step3: Access the work book */
			HSSFWorkbook wb = new HSSFWorkbook(xlDoc);

			/*Step4: Access the Sheet */
			HSSFSheet sheet = wb.getSheet(sheetName);

			int iRowCount = sheet.getLastRowNum() + 1;
			int iColCount =sheet.getRow(0).getLastCellNum();

			xlData = new String[iRowCount][iColCount];

			for(int i = 0; i < iRowCount; i ++)
			{
				for(int j = 0; j < iColCount; j ++)
				{
					HSSFCell cell = sheet.getRow(i).getCell(j);
					
					System.out.print(cell.getStringCellValue());
					System.out.print(" ");
					xlData[i][j] = cell.getStringCellValue();
				}
				System.out.println();
			}
			return xlData;
		}
	}


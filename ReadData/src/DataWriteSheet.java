
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.lang.invoke.SwitchPoint;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.CellStyle;


public class DataWriteSheet {

	public static void main(String[] args) throws IOException {

		String dataTablePath = "/Users/Fatema/Desktop/DataWriteSheet.xls";
		//String dataTablePath2 = "/Users/Fatema/Desktop/DataWriteSheet.xls";
		String sheetName = "Sheet1";
		String sheetName2 = "Sheet2";

		String recData[][] = null;
		String recData2[][] = null;
		recData = readSheet(dataTablePath, sheetName);
		recData2 = readSheet(dataTablePath, sheetName2);

		int iRowCount = recData.length;
		int iColumnCount = recData[0].length;


		String inputName;
		String run_NoRun;




		for(int i = 1; i < iRowCount ; i ++){
			run_NoRun = recData[i][1];

			if(run_NoRun.equalsIgnoreCase("Y")){
				/*inputName =recData[i][2];
				int flag = 0;
				switch (inputName) {
				case "Raj":
					System.out.println(inputName + " name is displayed in name1");
					break;
				case "Anusuya":
					System.out.println(inputName + " name is displayed in name2");
					break;
				case "Anu":
					System.out.println(inputName + " name is displayed in name3");
					break;
				case "Kumar":
					System.out.println(inputName + " name is displayed in name4");
					break;
				case "Bhavya":
					System.out.println(inputName + " name is displayed in name5");
					break;
				default:
					flag = 1;
					System.out.println(inputName + " name is not found in the system...");
					break;
				}
				if(flag==0){
					writeSheet(dataTablePath, sheetName, i, 3, "Present");
				}else{
					writeSheet(dataTablePath, sheetName, i, 3, "Not Present");
				}

			}else{
				writeSheet(dataTablePath, sheetName, i, 3, "No Run");*/
			}

		}

	}

	
	public static void writeSheet(String dataTablePath,String sheetName, int irow, int icolumn, String xlData) throws IOException{

		/* Access The Path of Xl File*/
		File xlFile = new File(dataTablePath );
		/* Access the File */
		FileInputStream xlDoc = new FileInputStream(xlFile);  
		/* Access The Work Book */
		HSSFWorkbook wb = new HSSFWorkbook(xlDoc);
		/* Access the Sheet */
		HSSFSheet sheet = wb.getSheet(sheetName);
		/*Access Row*/
		HSSFRow row = sheet.getRow(irow);
		/*Access Column*/
		HSSFCell cell = row.getCell(icolumn);
		/*Set cell for String value*/
		cell.setCellType(HSSFCell.CELL_TYPE_STRING);
		/*Enter the value*/
		cell.setCellValue(xlData);

		/* Set Color */
		HSSFCellStyle titleStyle = wb.createCellStyle();
		if(xlData.equalsIgnoreCase("Present")){
			titleStyle.setFillForegroundColor(new HSSFColor.GREEN().getIndex());

			titleStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);
			cell.setCellStyle(titleStyle);
		}else if(xlData.equalsIgnoreCase("Not Present")){
			titleStyle.setFillForegroundColor(new HSSFColor.RED().getIndex());
			titleStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);
			cell.setCellStyle(titleStyle);
		}else{
			titleStyle.setFillForegroundColor(new HSSFColor.WHITE().getIndex());
			titleStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);
			cell.setCellStyle(titleStyle);
		}



		FileOutputStream fout = new FileOutputStream(dataTablePath);
		wb.write(fout);
		fout.flush();
		fout.close();

	}

	/*	Method Name: readSheet
	 *  Method Description: Read data from Data Sheet
	 *  Arguments: dataTablePath = Path of the Data Sheet; sheetName = Name of Sheet
	 *  Created by: BSS 
	 *  Created Date: May 16 2014
	 *  Last Modified: May 16 2014 
	 */

	public static String[][] readSheet(String dataTablePath, String sheetName) throws IOException{
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

		for(int i = 0; i < iRowCount; i ++){
			for(int j = 0; j < iColCount; j ++){
				HSSFCell cell = sheet.getRow(i).getCell(j);
				xlData[i][j] = cell.getStringCellValue();
			}
		}
		return xlData;
	}
}

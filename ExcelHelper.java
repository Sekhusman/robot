package common;

import java.io.FileInputStream;
import java.util.ArrayList;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelHelper {

	private ArrayList<String> listheader;// 1D array
	private ArrayList<ArrayList<String>> listData;// 2d array

	public ExcelHelper() {
		listheader = new ArrayList<String>();
		listData = new ArrayList<ArrayList<String>>();

	}

	public void setListHeader(String filename, int headerindex) {
		try {

			// ist create a test excel and convert to text
			FileInputStream fis = new FileInputStream(filename);
			XSSFWorkbook wb = new XSSFWorkbook(fis);
			XSSFSheet sheet = wb.getSheet("Sheet1");
			XSSFRow headerrow = sheet.getRow(headerindex);
			for (int i = 0; i < headerrow.getLastCellNum(); i++) {

				XSSFCell cell = headerrow.getCell(i);

				if (cell == null) {
					listheader.add("");
					System.out.println("");
				} else {
					listheader.add(String.valueOf(cell));
					// System.out.println(String.valueOf(cell));

				}
				System.out.println(listheader);

				wb.close();
				fis.close();
			}

		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	public void setListData(String filename, String tcname) {
		try {

			FileInputStream fis = new FileInputStream(filename);
			XSSFWorkbook wb = new XSSFWorkbook(fis);
			XSSFSheet sheet = wb.getSheet("Sheet1");
			for (int i = 0; i <= sheet.getLastRowNum(); i++) {
				XSSFRow dataRow = sheet.getRow(i);
				String xlTcname = String.valueOf(dataRow.getCell(0));
				if (xlTcname.equalsIgnoreCase(tcname)) {

					ArrayList<String> tempdata = new ArrayList<String>();
					for (int j = 0; j < dataRow.getLastCellNum(); j++) {

						XSSFCell cell = dataRow.getCell(j);

						if (cell == null) {
							tempdata.add("");
							// System.out.println("");
						} else {
							tempdata.add(String.valueOf(cell));
							// System.out.println(String.valueOf(cell));

						}

					}
					listData.add(tempdata);

				}

			}
			System.out.println(listData);
			wb.close();
			fis.close();
		} catch (Exception e) {
			e.printStackTrace();

		}

	}

	public String  getValue(int rowindex,String columnname) {
	String value="No value"	;
	try{
		int colindex=listheader.indexOf(columnname);
		value=listData.get(rowindex).get(colindex);
		
		//System.out.println(arg0);
	}catch(Exception e)
		
	{
		e.printStackTrace();
	}
		
return value;
	}

	
	public void clearlist()
	
	{
		listData.clear();
	}
	
	
}

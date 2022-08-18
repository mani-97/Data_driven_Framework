import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.NumberToTextConverter;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class DrivernTest {

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub


		FileInputStream file=new FileInputStream("D:\\Excel.xlsx");

		XSSFWorkbook workbook=new XSSFWorkbook(file);

		List<String> list=new ArrayList<String>();

		int Sheetcount=workbook.getNumberOfSheets();
		for(int i=0;i<Sheetcount;i++) {
			if(workbook.getSheetName(i).equalsIgnoreCase("Sheet1")) {	
				XSSFSheet sheet=workbook.getSheetAt(i);
				Iterator<Row> rows=sheet.iterator();
				Row firstrow=rows.next();
				Iterator<Cell> cells=firstrow.iterator();
				int k=0;
				int col=0;
				while(cells.hasNext()) {
					if(cells.next().getStringCellValue().equalsIgnoreCase("word")) {
						k=col;
					}
					k++;
				}
				while(rows.hasNext()) {
					Row r=rows.next();	
					if(r.getCell(col).getStringCellValue().equalsIgnoreCase("Excel")) {	
						Iterator<Cell>cell=r.cellIterator();
						while(cell.hasNext()) {
							Cell celvalue=cell.next();
							if(celvalue.getCellType()==CellType.STRING) {
								list.add(celvalue.getStringCellValue());	
							}
							else
							{
								String numValue=NumberToTextConverter.toText(celvalue.getNumericCellValue());
								list.add(numValue);								
							}	
						}
					}
				}		
			}
		}


		System.out.println(list);














	}
}

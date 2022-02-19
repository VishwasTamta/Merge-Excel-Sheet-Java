package merge;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.HashMap;
import java.util.Map;
import java.util.Set;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;  

public class MergeFiles {
	public static void main(String[] args) {
		
	try {
		// Fetching Files
		FileInputStream f1 = new FileInputStream(new File("C:\\Users\\Vishwas\\eclipse-workspace\\merge\\file1.xlsx"));
		FileInputStream f2 = new FileInputStream(new File("C:\\Users\\Vishwas\\eclipse-workspace\\merge\\file3.xlsx"));
		
		FileInputStream f3 = new FileInputStream(new File("C:\\Users\\Vishwas\\eclipse-workspace\\merge\\file2.xlsx"));
		FileInputStream f4 = new FileInputStream(new File("C:\\Users\\Vishwas\\eclipse-workspace\\merge\\file4.xlsx"));
		
		//Workbook
		XSSFWorkbook wb1 = new XSSFWorkbook(f1);
		XSSFWorkbook wb2 = new XSSFWorkbook(f2);
		
		XSSFWorkbook wb3 = new XSSFWorkbook(f3);
		XSSFWorkbook wb4 = new XSSFWorkbook(f4);
		
		//Sheets
		XSSFSheet s1 = wb1.getSheetAt(0);
		XSSFSheet s2 = wb2.getSheetAt(0);
		XSSFSheet s3 = wb3.getSheetAt(0);
		XSSFSheet s4 = wb4.getSheetAt(0);
		
		//Adding sheets
		addSheet(s1,s2,mapHeaders(s2,s1));
		addSheet(s3,s4,mapHeaders(s4,s3));
		f1.close();
		f2.close();
		f3.close();
		f4.close();
		
		// save file
		File newFile = new File("C:\\Users\\Vishwas\\eclipse-workspace\\merge\\file1-file3.xlsx");
		
		File newFile1 = new File("C:\\Users\\Vishwas\\eclipse-workspace\\merge\\file2-file4.xlsx");
		
		if(!newFile.exists()) {
			newFile.createNewFile();
		}
		
		if(!newFile1.exists()) {
			newFile1.createNewFile();
		}
		
		FileOutputStream out = new FileOutputStream(newFile);
		
		FileOutputStream out1 = new FileOutputStream(newFile1);
		
		wb1.write(out);
		wb3.write(out1);
		
		out.close();
		out1.close();
		System.out.println("File merged successfully");
	} catch(Exception e) {
		e.printStackTrace();
	}
}



public static void addSheet(XSSFSheet mainSheet, XSSFSheet sheet, HashMap<Integer, Integer> map) {
	
	Set<Integer> colNum = map.keySet();
	Map<Integer, XSSFCellStyle> styleMap = new HashMap<Integer, XSSFCellStyle>();
	
	int len = mainSheet.getLastRowNum();
	for(int j=sheet.getFirstRowNum() + 1; j<= sheet.getLastRowNum(); j++) {
		XSSFRow row = sheet.getRow(j);
		XSSFRow mRow = mainSheet.createRow(len + j);
		
		for(Integer k : colNum) {
			XSSFCell cell = row.getCell(k.intValue());
			XSSFCell mcell = mRow.createCell(map.get(k).intValue());
			if(cell.getSheet().getWorkbook() == mcell.getSheet().getWorkbook()) {
				mcell.setCellStyle(cell.getCellStyle());
			}else {
				int stHashCode = cell.getCellStyle().hashCode();
				XSSFCellStyle newCellStyle = styleMap.get(stHashCode);
				if(newCellStyle == null) {
					newCellStyle = mcell.getSheet().getWorkbook().createCellStyle();
					newCellStyle.cloneStyleFrom(cell.getCellStyle());
					styleMap.put(stHashCode, newCellStyle);
				}
				mcell.setCellStyle(newCellStyle);
			}
			
			switch(cell.getCellType()) {
			case HSSFCell.CELL_TYPE_FORMULA:
				mcell.setCellFormula(cell.getCellFormula());
				break;
			case HSSFCell.CELL_TYPE_NUMERIC:
				mcell.setCellValue(cell.getNumericCellValue());
				break;
			case HSSFCell.CELL_TYPE_STRING:
				mcell.setCellValue(cell.getStringCellValue());
				break;
			case HSSFCell.CELL_TYPE_BLANK:
				mcell.setCellType(HSSFCell.CELL_TYPE_BLANK);
				break;
			case HSSFCell.CELL_TYPE_BOOLEAN:
				mcell.setCellValue(cell.getBooleanCellValue());
				break;
			case HSSFCell.CELL_TYPE_ERROR:
				mcell.setCellErrorValue(cell.getErrorCellValue());
				break;
			default:
				mcell.setCellValue(cell.getStringCellValue());
				break;
			}
		}
	}
}

public static HashMap<Integer, Integer> mapHeaders(XSSFSheet sheet1,
        XSSFSheet sheet2) {
    HashMap<Integer, Integer> map = new HashMap<Integer, Integer>();
    XSSFRow row1 = sheet1.getRow(0);
    XSSFRow row2 = sheet2.getRow(0);
    for (int i = row1.getFirstCellNum(); i < row1.getLastCellNum(); i++) {
        for (int j = row2.getFirstCellNum(); j < row2.getLastCellNum(); j++) {
            if (row1.getCell(i).getStringCellValue()
                    .equals(row2.getCell(j).getStringCellValue())) {
                map.put(new Integer(i), new Integer(j));
            }
        }
    }
    return map;
}
}



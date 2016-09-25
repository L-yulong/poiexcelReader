package org.rain.poireadutill;

import java.io.InputStream;
import java.util.List;

public interface ExcelReader {
	 public List<String[]> getSheet(int page);
	 
	 public int getSheetNum();
	 
	 public void readExcelContent(InputStream in,int rowNum);
	 
	 public List<String[]> readSheetContent(InputStream in, int rowNum, int sheetNum);
}

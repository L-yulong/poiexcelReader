package org.rain.poireadutill;

public class ExcelReaderHelp {
	public ExcelReader getExcelRead03(){
		ExcelReader2003 excelReader2003 = new ExcelReader2003();
		return (ExcelReader)excelReader2003;
	}
	
	public ExcelReader getExcelRead07(){
		ExcelReader2007 excelReader2007 = new ExcelReader2007();
		return (ExcelReader)excelReader2007;
	}
}

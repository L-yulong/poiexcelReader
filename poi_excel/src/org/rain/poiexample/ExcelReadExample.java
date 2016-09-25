package org.rain.poiexample;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.InputStream;
import java.util.List;

import org.rain.poireadutill.ExcelReader;
import org.rain.poireadutill.ExcelReaderHelp;

public class ExcelReadExample {
	
	/**
	 * 读取03版Excel
	 * @throws FileNotFoundException
	 */
	public void read03() throws FileNotFoundException{
		File excelFile = new File(ExcelReadExample.class.getResource("03Excel.xls").getFile());
		InputStream inputStream = new FileInputStream(excelFile);
		
		// 工具读取Excel
		ExcelReaderHelp excelReaderHelp =  new ExcelReaderHelp();
		ExcelReader excelReader = excelReaderHelp.getExcelRead03();
		excelReader.readExcelContent(inputStream, 2);	
		
		List<String[]>  list = excelReader.getSheet(0);	
		System.out.println("打印03Excel");
		printList(list);
	}
	
	/**
	 * 读取07版Excel
	 * @throws FileNotFoundException
	 */
	public void read07() throws FileNotFoundException{
		File excelFile = new File(ExcelReadExample.class.getResource("07Excel.xlsx").getFile());
		InputStream inputStream = new FileInputStream(excelFile);
		// 工具读取Excel
		ExcelReaderHelp excelReaderHelp =  new ExcelReaderHelp();
		ExcelReader excelReader = excelReaderHelp.getExcelRead07();
		excelReader.readExcelContent(inputStream, 2);	
		
		List<String[]>  list = excelReader.getSheet(0);	
		System.out.println("打印07Excel");
		printList(list);
	}
	
	/**
	 * 打印List
	 * @param list
	 */
	public void printList(List<String[]> list){
		for(int i = 0, size = list.size(); i < size; i++){
			System.out.println("row" + i);
			String[] str = list.get(i);
			for(int j = 0, length = str.length; j < length; j++){
				System.out.print(str[j] +", ");
			}
			System.out.println("");
		}	
	}
	
	public  static void main(String[] arg) throws FileNotFoundException{		
		ExcelReadExample excelReadExample = new ExcelReadExample();
		excelReadExample.read03();
		excelReadExample.read07();
	}
}

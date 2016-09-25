package org.rain.poireadutill;

import java.io.BufferedInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

import javax.xml.parsers.SAXParser;
import javax.xml.parsers.SAXParserFactory;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xssf.eventusermodel.ReadOnlySharedStringsTable;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.model.StylesTable;
import org.xml.sax.InputSource;
import org.xml.sax.XMLReader;

public class ExcelReader2007 implements ExcelReader{
	//记录每个sheet数据
	public List<List<String[]>> listforpage = new ArrayList<List<String[]>>();
	private int sheetNo=0;
	//行大小
	private int defaultrowNum = 28;
	//是否只读取一个sheet
	private boolean isReadOnlySheet = false;
	//只读sheet的序号
	private int onlySheetNum;
	
	public void readExcelContent(InputStream in) {
		readExcelContent(in,defaultrowNum);
	}
	
	/**
	 * 读取一个sheet的数据
	 * @param in
	 * @param rowNum
	 * @param sheetNum
	 * @return
	 */
	public List<String[]> readSheetContent(InputStream in, int rowNum, int sheetNum){
		isReadOnlySheet = true;
		onlySheetNum = sheetNum;
    	if(listforpage.size()>sheetNum)
    		return listforpage.get(0);
    	else 
    		return null;  	
	}
    
    /**
     * 读取Excel数据内容
     * @param InputStream
     * @return Map 包含单元格数据内容的Map对象
     */
    public void readExcelContent(InputStream in,int rowNum) {
    	BufferedInputStream bis = null;
    	try{
    		bis = new BufferedInputStream(in);
    		//取得excel的poi解析对象
    		OPCPackage p = OPCPackage.open(bis);
    		ReadOnlySharedStringsTable strings = new ReadOnlySharedStringsTable(p);  		
    		XSSFReader xssfReader = new XSSFReader(p);
    		StylesTable styles = xssfReader.getStylesTable();
            XSSFReader.SheetIterator iter = (XSSFReader.SheetIterator) xssfReader.getSheetsData();
            // 遍历excel的sheet
            while (iter.hasNext()) {
            	if(isReadOnlySheet){
            		if(sheetNo != onlySheetNum) continue;
            	}
            	//读取每个sheet的流,使用sax方式进行解析，通过getRows()返回
                InputStream stream = iter.next();  
                InputSource sheetSource = new InputSource(stream);
                ReadExcel2007ByPoi handler = new ReadExcel2007ByPoi(styles,strings,rowNum, System.out);
                SAXParserFactory saxFactory = SAXParserFactory.newInstance();
                SAXParser saxParser = saxFactory.newSAXParser();
                XMLReader sheetParser = saxParser.getXMLReader();
                sheetParser.setContentHandler(handler);
                sheetParser.parse(sheetSource);
                List<String[]> list = handler.getRows();
    	        listforpage.add(list);
    	        sheetNo++;
                stream.close(); 
            }           
    	}catch(Exception e){
    		throw new RuntimeException(e);
    	}finally{  		
            if (bis != null) {  
                try {  
                	bis.close();  
                } catch (IOException io) {  
                	throw new RuntimeException(io);
                }  
            } 
    	}
    }
    
    /**
     * 获取sheet的数据
     */
    public List<String[]> getSheet(int page){
    	if(listforpage.size()>page)
    		return listforpage.get(page);
    	else 
    		return null;  		
    }
    
    /**
     * 获取sheet数量
     */
    public int getSheetNum(){
    	return sheetNo;
    }
}

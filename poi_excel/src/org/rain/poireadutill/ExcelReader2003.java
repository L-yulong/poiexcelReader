package org.rain.poireadutill;

import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.hssf.eventusermodel.FormatTrackingHSSFListener;
import org.apache.poi.hssf.eventusermodel.HSSFEventFactory;
import org.apache.poi.hssf.eventusermodel.HSSFRequest;
import org.apache.poi.hssf.eventusermodel.MissingRecordAwareHSSFListener;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

public class ExcelReader2003 implements ExcelReader{
	//记录每个sheet数据
	public List<List<String[]>> listforpage = new ArrayList<List<String[]>>();
	private int sheetNo=0;
	//行大小
	private int defaultrowNum = 28;
	//
	private ReadExcel2003ByPoi readExcel2003ByPoi = null;
	//
	private FormatTrackingHSSFListener formatListener = null;
	
	public void readExcelContent(InputStream in) {
		readExcelContent(in,defaultrowNum);
	}
    
    /**
     * 读取Excel数据内容
     * @param InputStream
     * @return Map 包含单元格数据内容的Map对象
     */
    public void readExcelContent(InputStream in,int rowNum) {
		POIFSFileSystem poifs = null;
		InputStream din = null;
		try{
			poifs = new POIFSFileSystem(in);  
			din = poifs.createDocumentInputStream("Workbook");  
	        //这儿为所有类型的Record都注册了监听器，如果需求明确的话，可以用addListener方法，并指定所需的Record类型     
	        HSSFRequest req = new HSSFRequest();  
	        //添加监听记录的事件   
	        readExcel2003ByPoi = new ReadExcel2003ByPoi();
	        readExcel2003ByPoi.setRowNum(rowNum);
	        MissingRecordAwareHSSFListener listener = new MissingRecordAwareHSSFListener(readExcel2003ByPoi); 
	        //监听代理，方便获取record format
	        formatListener = new FormatTrackingHSSFListener(listener);    
	        req.addListenerForAllRecords(formatListener);
	        //创建事件工厂  
	        HSSFEventFactory factory = new HSSFEventFactory();  
	        //处理基于时间文档流(循环获取每一条Record进行处理)  
	        factory.processEvents(req, din); 
		}catch(Exception e){
			throw new RuntimeException(e);
		}finally{            
	        //关闭基于POI文档流
            if (din != null) {  
                try {  
                	din.close();  
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
    	return readExcel2003ByPoi.getSheet(page); 		
    }
    
	/**
	 * 读取一个sheet的数据
	 * @param in
	 * @param rowNum
	 * @param sheetNum
	 * @return
	 */
	public List<String[]> readSheetContent(InputStream in, int rowNum, int sheetNum){
		readExcelContent(in, rowNum);
		return readExcel2003ByPoi.getSheet(sheetNum);	
	}
    
    /**
     * 获取sheet数量
     */
    public int getSheetNum(){
    	return sheetNo;
    }
}

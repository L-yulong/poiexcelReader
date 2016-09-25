package org.rain.poireadutill;

import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.hssf.eventusermodel.HSSFListener;
import org.apache.poi.hssf.eventusermodel.dummyrecord.LastCellOfRowDummyRecord;
import org.apache.poi.hssf.record.BOFRecord;
import org.apache.poi.hssf.record.BlankRecord;
import org.apache.poi.hssf.record.BoolErrRecord;
import org.apache.poi.hssf.record.BoundSheetRecord;
import org.apache.poi.hssf.record.FormulaRecord;
import org.apache.poi.hssf.record.LabelSSTRecord;
import org.apache.poi.hssf.record.NumberRecord;
import org.apache.poi.hssf.record.Record;
import org.apache.poi.hssf.record.RowRecord;
import org.apache.poi.hssf.record.SSTRecord;
import org.apache.poi.hssf.record.StringRecord;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;

public class ReadExcel2003ByPoi implements HSSFListener{
	/** 默认读取列数 **/
	private int rowNum = 28;
	private Map<Integer,String> formatMap = new HashMap<Integer,String>();
	/**　记录每个sheet数据　**/
	public List<List<String[]>> listforpage = new ArrayList<List<String[]>>();
	/** sheet记录 **/
	List<String[]> list = null;
	/**　row记录　**/
	String[] rowArr = null;
    /** 当前列 **/  
    private int curRowNum=0; 
    /** excel的sheet的数量 **/
    private int sheetNo=0;  
	private SSTRecord sstrec;
    
	/**
	 * 设置读取的列数
	 */
	public void setRowNum(int rowNum){
		this.rowNum = rowNum;
	}
	
	/**
	 * 设置列格式
	 * @param index
	 * @param type
	 */
	public void setRowFormat(int index, String type){
		formatMap.put(index, type);
	}
	
    /**
     * 监控入口 
     */
    public void processRecord(Record record) { 
    	DecimalFormat df = new DecimalFormat("0.00");
    	
        switch (record.getSid()) {  
          
        case BOFRecord.sid:  
            BOFRecord bof = (BOFRecord) record;  
            //顺序进入新的Workbook    
            if (bof.getType() == bof.TYPE_WORKBOOK) {  
            	
            //顺序进入新的sheet页  
            } else if (bof.getType() == bof.TYPE_WORKSHEET) {              	
            	if(sheetNo>0)
            	listforpage.add(list);
            	list = new ArrayList<String[]>();
            	rowArr = new String[rowNum];
            	sheetNo++;
            }  
            break;  
        //开始解析Sheet的信息,获取sheet的名称等信息     
        case BoundSheetRecord.sid:  
            BoundSheetRecord bsr = (BoundSheetRecord) record; 
            break;  
        //执行行记录事件  
        case RowRecord.sid:
        	RowRecord rowrec = (RowRecord) record; 
            break;  
        // SSTRecords store a array of unique strings used in Excel.  
        case SSTRecord.sid: 
        	sstrec = (SSTRecord) record;
            break;               
        //发现数字类型的cell     
        case NumberRecord.sid:       	
        	NumberRecord nr = (NumberRecord) record;
        	curRowNum = nr.getColumn();           
    		String rowFormmat = formatMap.get(curRowNum);
        	 
        	String valueNum =  df.format(nr.getValue());
        	//excel中指定分类为日期格式
        	if(rowFormmat != null && rowFormmat.equalsIgnoreCase("Date")){
				String timeStr=(new SimpleDateFormat("yyyy-MM-dd")).format(HSSFDateUtil.getJavaDate(nr.getValue()));
				rowArr[curRowNum] = timeStr;
        	}else{
        		rowArr[curRowNum] = valueNum + "";
        	}
            break;  
        case FormulaRecord.sid: //单元格为公式类型  
            FormulaRecord frec = (FormulaRecord) record;  
            int curRowNumF = frec.getColumn();
                if (!Double.isNaN(frec.getValue())) {  
                	rowArr[curRowNumF] =  df.format(frec.getValue());
                }

            break;  
        case StringRecord.sid://单元格中公式的字符串  
                StringRecord srec = (StringRecord) record; 
                rowArr[curRowNum] =  srec.getString();

            break; 
        //发现字符串类型，这儿要取字符串的值的话，跟据其index去字符串表里读取     
        case LabelSSTRecord.sid:       	
            LabelSSTRecord lsr = (LabelSSTRecord)record; 
            curRowNum = lsr.getColumn();
            rowArr[curRowNum] =  sstrec.getString(lsr.getSSTIndex()).toString().trim();           
            break;  
        case BoolErrRecord.sid: //解析boolean错误信息        	
            BoolErrRecord ber = (BoolErrRecord)record;
        	curRowNum = ber.getColumn();     
            rowArr[curRowNum] =  ber.getBooleanValue()+"";
            break;     
         //空白记录的信息  
        case BlankRecord.sid: 
        	BlankRecord bla = (BlankRecord)record;
        	curRowNum = bla.getColumn();
        	rowArr[curRowNum] = "";     
            break;        
        }               
        // 行结束时的操作  
        if (record instanceof LastCellOfRowDummyRecord) {
            if(curRowNum!=0&&checkNullRow(rowArr)){
            	list.add(rowArr);
            	curRowNum = 0;
            }
        	rowArr = new String[rowNum];
        }  
        
    }
    
    /**
     * 获取sheet的数据
     */
    public List<String[]> getSheet(int page){
    	if(listforpage.size()>page)
    		return listforpage.get(page);
    	else 
    		return list;  		
    }
    
    /**
     * 获取sheet数量
     */
    public int getSheetNum(){
    	return sheetNo;
    }
    
    /**
     * 检查是否为空行
     * @param obj
     * @return
     */
    public boolean checkNullRow(Object[] obj){
    	boolean bl = false;
    	String temp;
    	for(int i=0,size = obj.length;i<size;i++){
    		temp = (String)obj[i];
    		if(temp==null||temp.trim().length()==0)continue;
    		bl = true;
    		break;
    	}  	
    	return bl;
    }
}  
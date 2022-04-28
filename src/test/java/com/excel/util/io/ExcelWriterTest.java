package com.excel.util.io;

import static org.junit.Assert.assertTrue;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.junit.Test;

import com.excel.util.model.SheetModel;

/**
 * Unit Test Class for ExcelWriter
 *
 * @author yukesh
 *
 */
public class ExcelWriterTest {

    @Test
    public void writeToExcel() {

        
    	String fileAbsPath = ExcelWriterTest.class.getResource("/").getPath();
        boolean result = ExcelWriter.getInstance().writeToExcel(mockSheetModel(), fileAbsPath + "writeToExcel.xlsx");
        assertTrue(result);
    }
    
    @Test
    public void writeToExcelUsingMap() {
    	
    	String fileAbsPath = ExcelWriterTest.class.getResource("/").getPath();
        boolean result = ExcelWriter.getInstance().writeToExcel(mockRowHeaderMapList(), "Test", fileAbsPath + "writeToExcelUsingMap.xlsx");
        assertTrue(result);
    }

    public SheetModel mockSheetModel(){
        SheetModel sheetModel = new SheetModel();

        sheetModel.setSheetName("Test Sheet");
        sheetModel.setHeaderAttrs(new String[] {"Col1", "Col2", "Col3"});
        sheetModel.getRowValueArr().add(new String[]{"Row1", "Row2", "Row3"});

        return sheetModel;
    }
    
    public List<Map<String, String>> mockRowHeaderMapList(){
    	List<Map<String,String>> rowHeaderDataMapList = new ArrayList<Map<String, String>>();
    	
    	Map<String, String> headerDataMap = new HashMap<String, String>();
    	headerDataMap.put("Col1","Row1 Cell1");
    	headerDataMap.put("Col2","Row1 Cell2");
    	headerDataMap.put("Col3","Row1 Cell3");
    	
    	rowHeaderDataMapList.add(headerDataMap);
    	
    	return rowHeaderDataMapList;
    }

}

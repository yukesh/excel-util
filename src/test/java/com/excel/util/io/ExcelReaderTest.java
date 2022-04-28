package com.excel.util.io;

import static org.junit.Assert.assertTrue;

import java.io.FileNotFoundException;
import java.net.URL;
import java.util.List;
import java.util.Map;

import org.junit.Test;

public class ExcelReaderTest {


    @Test
    public void readExcel() throws Exception {
    	
    	String fileName = null;
    	ClassLoader classLoader = ExcelReaderTest.class.getClassLoader();
    	URL url = classLoader.getResource("TestFile.xlsx");
    	if(null != url) {
    		fileName = url.getPath();
    	}
    	System.out.println("File Name :- " + fileName);
    	
        ExcelReader excelReader = new ExcelReader(fileName);
        List<Map<String,String>> rowHeaderDataMapList = excelReader.getRowHeaderDataMapList();
        
        assert(!rowHeaderDataMapList.isEmpty());

    }

    @Test
    public void fileNotFoundExcel() {
    	
    	String fileName = "Test.xlsx";
    	
    	ExcelReader excelReader;
		try {
			excelReader = new ExcelReader(fileName);
			excelReader.getRowHeaderDataMapList();
		} catch (Exception exception) {
			assertTrue(exception instanceof FileNotFoundException);
		}
    }

}

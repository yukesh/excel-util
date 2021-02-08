package com.excel.util.io;

import org.junit.Test;

import java.util.List;
import java.util.Map;

public class ExcelReaderTest {


    @Test
    public void readExcel(){

        String fileName = "/Users/yukesh/Downloads/Test-Book.xlsx";

        ExcelReader excelReader = new ExcelReader(fileName);
        List<Map<String,String>> rowHeaderDataMapList = excelReader.getRowHeaderDataMapList();

        assert(!rowHeaderDataMapList.isEmpty());

    }


}

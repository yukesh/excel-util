package com.excel.util.io;

import com.excel.util.helper.FileReadWriteHelper;
import com.excel.util.model.SheetModel;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.util.ArrayList;
import java.util.List;

/**
 *
 * JAVA Class for Writing Excel File
 *
 * @author yukesh
 */
public class ExcelWriter {

    /**
     * Private Constructor to avoid instance creation from outside
     */
    private ExcelWriter(){
        super();
    }

    /**
     * Inner Static Class to return the Singleton instance
     */
    private static class ExcelWriterSingleton {
        private static final ExcelWriter INSTANCE = new ExcelWriter();
    }

    /**
     * Static method to return the singleton instance
     * @return ExcelWriter
     */
    public static ExcelWriter getInstance(){
        return ExcelWriterSingleton.INSTANCE;
    }

    public boolean writeToExcel(String aFilePath, SheetModel aSheetModel) {

        boolean isSuccess = false;
        if(StringUtils.isNotBlank(aFilePath) && null != aSheetModel){

            String sheetName = aSheetModel.getSheetName();

            XSSFWorkbook xssfWorkbook = new XSSFWorkbook();
            XSSFSheet xssfSheet = StringUtils.isNotBlank(sheetName)
                    ? xssfWorkbook.createSheet(sheetName) : xssfWorkbook.createSheet();

            String[] headerAttrs = aSheetModel.getHeaderAttrs();
            List<String[]> rowValueArr = new ArrayList<>();
            if(null != headerAttrs){
                rowValueArr.add(headerAttrs);
            }
            rowValueArr.addAll(aSheetModel.getRowValueArr());

            for(int rowNum = 0; rowNum < rowValueArr.size(); rowNum++){

                Row row = xssfSheet.createRow(rowNum);
                String[] rowAttrs = rowValueArr.get(rowNum);

                for(int cellNum = 0; cellNum < rowAttrs.length; cellNum++) {
                    Cell cell = row.createCell(cellNum);
                    cell.setCellValue(rowAttrs[cellNum]);
                }
            }
            //Write the Workbook to the File
            isSuccess = FileReadWriteHelper.writeToFile(aFilePath, xssfWorkbook);
        }
        return isSuccess;
    }

}

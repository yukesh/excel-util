package com.excel.util.io;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.Set;

import org.apache.commons.lang3.StringUtils;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.excel.util.model.SheetModel;

/**
 *
 * JAVA Class for Writing Excel File
 *
 * @author yukesh
 */
public class ExcelWriter {
	
	private static final Logger logger = LogManager.getLogger(ExcelWriter.class);
	
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

    /**
     * Method to write the SheetModel to file
     * @param aSheetModel
     * @param aFilePath
     * @return
     */
    public boolean writeToExcel(SheetModel aSheetModel, String aFilePath) {

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
            isSuccess = writeToFile(xssfWorkbook, aFilePath);
        }
        return isSuccess;
    }
    
    /**
     * Method to Write the XSSFWorkbook to the file.
     * @param aFilePath
     * @param aXssfWorkbook
     * @return
     */
    public boolean writeToFile(XSSFWorkbook aXssfWorkbook, String aFilePath){
        boolean isSuccess = false;

        if(null != aXssfWorkbook){
        	
            try {

                FileOutputStream fileOutputStream = new FileOutputStream(aFilePath);
                aXssfWorkbook.write(fileOutputStream);
                fileOutputStream.close();
                isSuccess = true;
                logger.info("Workbook has been created Successfully! - [" + aFilePath +"]");
            } catch (FileNotFoundException fileNotFoundException) {
            	logger.warn("File Not Found Exception Occurred while writing to [" + aFilePath +"]");
            	logger.error(fileNotFoundException);

            } catch (IOException ioException) {
            	logger.warn("IO Exception Occurred while writing to [" + aFilePath +"]");
            	logger.error(ioException);
            }
        }

        return isSuccess;
    }
    
    /**
	 * Method to convert the RowHeaderDataMap to SheetModel and write to Excel workbook.
	 * 
	 * @param aRowHeaderDataMapList
	 * @param aSheetName
	 * @param aFilePath
	 */
	public boolean writeToExcel(List<Map<String, String>> aRowHeaderDataMapList, String aSheetName, String aFilePath) {

		boolean isSuccess = false;

		if (!aRowHeaderDataMapList.isEmpty()) {
			
			String sheetName = StringUtils.isNotBlank(aSheetName) ? aSheetName : "Sheet";
			SheetModel sheetModel = new SheetModel(sheetName);

			Set<String> headerSet = aRowHeaderDataMapList.get(0).keySet();
			sheetModel.setHeaderAttrs(headerSet.toArray(new String[headerSet.size()]));

			aRowHeaderDataMapList.forEach(headerDataMap -> {

				List<String> rowValueList = new ArrayList<>();
				headerSet.forEach(header -> {
					rowValueList.add(headerDataMap.get(header));
				});

				sheetModel.getRowValueArr().add(rowValueList.toArray(new String[rowValueList.size()]));

			});

			isSuccess = writeToExcel(sheetModel, aFilePath);
		}

		return isSuccess;

	}

}

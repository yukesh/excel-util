package com.excel.util.helper;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

public class FileReadWriteHelper {

    /**
     * Helper Class to Write the XSSFWorkbook to the file.
     * @param aFilePath
     * @param aXssfWorkbook
     * @return
     */
    public static boolean writeToFile(String aFilePath, XSSFWorkbook aXssfWorkbook){
        boolean isSuccess = false;

        if(null != aXssfWorkbook){

            File outputFile = new File(aFilePath);
            try {

                FileOutputStream fileOutputStream = new FileOutputStream(aFilePath);
                aXssfWorkbook.write(fileOutputStream);
                fileOutputStream.close();
                isSuccess = true;
                System.out.println("Workbook has been created Successfully! - [" + aFilePath +"]");
            } catch (FileNotFoundException e) {
                System.out.println("File Not Found Exception Occurred while writing to [" + aFilePath +"]");
                e.printStackTrace();

            } catch (IOException e) {
                System.out.println("IO Exception Occurred while writing to [" + aFilePath +"]");
                e.printStackTrace();
            }
        }

        return isSuccess;
    }

}

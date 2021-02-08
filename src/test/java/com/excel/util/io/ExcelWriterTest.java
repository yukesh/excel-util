package com.excel.util.io;

import com.excel.util.io.ExcelWriter;
import com.excel.util.model.SheetModel;
import org.junit.Test;

import static org.junit.Assert.assertTrue;

/**
 * Unit Test Class for ExcelWriter
 *
 * @author yukesh
 *
 */
public class ExcelWriterTest {

    @Test
    public void writeToExcel() {

        String fileAbsPath = "/Users/yukesh/Workspaces/Output.xlsx";
        boolean result = ExcelWriter.getInstance().writeToExcel(fileAbsPath, mockSheetModel());
        assertTrue(result);
    }

    public SheetModel mockSheetModel(){
        SheetModel sheetModel = new SheetModel();

        sheetModel.setSheetName("Test Sheet");
        sheetModel.setHeaderAttrs(new String[] {"Col1", "Col2", "Col3"});
        sheetModel.getRowValueArr().add(new String[]{"Row1", "Row2", "Row3"});

        return sheetModel;
    }

}

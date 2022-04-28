package com.excel.util.model;

import java.util.ArrayList;
import java.util.List;

/**
 * JAVA Model Class for Excel Sheet
 *
 * @author yukesh
 *
 */
public class SheetModel {

    private String sheetName;
    private String[] headerAttrs;
    private List<String[]> rowValueArr;

    public SheetModel() {
    	super();
    }
    
    public SheetModel(String aSheetName) {
    	sheetName = aSheetName;
    }
    
    /**
     * Set the Sheet Name
     * @param aSheetName
     */
    public void setSheetName(String aSheetName){
        sheetName = aSheetName;
    }

    /**
     * Return the Sheet Name
     * @return
     */
    public String getSheetName(){
        return sheetName;
    }

    /**
     * Set the Sheet Header Attributes
     * @param aHeaderAttrs
     */
    public void setHeaderAttrs(String[] aHeaderAttrs){
        headerAttrs = aHeaderAttrs;
    }

    /**
     * Return the Sheet Header Attributes
     * @return
     */
    public String[] getHeaderAttrs(){
        return headerAttrs;
    }

    /**
     * Return the List of Sheet Row Value Attributes
     * @return
     */
    public List<String[]> getRowValueArr(){

        if(null == rowValueArr){
            rowValueArr = new ArrayList<String[]>();
        }
        return rowValueArr;
    }
}

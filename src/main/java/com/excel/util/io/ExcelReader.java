package com.excel.util.io;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

import javax.xml.stream.XMLInputFactory;
import javax.xml.stream.XMLStreamException;
import javax.xml.stream.XMLStreamReader;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackageAccess;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.eventusermodel.ReadOnlySharedStringsTable;
import org.apache.poi.xssf.eventusermodel.XSSFReader;

/**
 *
 * JAVA Class for Reading Excel File
 *
 * @author yukesh
 */
public class ExcelReader {

    private ReadOnlySharedStringsTable stringsTable;
    private XSSFReader xssfReader;
    private XMLInputFactory inputFactory;
    private OPCPackage excelPackage;

    private static final String ELEMENT_ROW = "row";
    private static final String ELEMENT_C = "c";
    private static final String ELEMENT_R = "r";
    private static final String ELEMENT_V = "v";
    private static final String ELEMENT_S = "s";
    private static final String ELEMENT_T = "t";

    private List<String> sheetNames;

    private void setSheetNames() throws InvalidFormatException, IOException {

        XSSFReader.SheetIterator sheetItr = (XSSFReader.SheetIterator) xssfReader.getSheetsData();
        while(sheetItr.hasNext()) {
            sheetItr.next();
            String sheetName = sheetItr.getSheetName();
            getSheetNames().add(sheetName);
        }
    }

    public List<String> getSheetNames(){
        if(null == sheetNames) {
            sheetNames = new ArrayList<>();
        }
        return sheetNames;
    }

    /**
     * @param aFilePath - Instance using the File location
     */
    public ExcelReader(String aFilePath) throws Exception {
        this(new File(aFilePath));
    }

    /**
     * @param aFile - Instance using the Input file
     * @throws InvalidFormatException 
     */
    public ExcelReader(File aFile) throws Exception {
    	if(aFile.isFile()) {
            excelPackage = OPCPackage.open(aFile, PackageAccess.READ);
            stringsTable = new ReadOnlySharedStringsTable(excelPackage);
            xssfReader = new XSSFReader(excelPackage);
            inputFactory = XMLInputFactory.newInstance();

            setSheetNames();

        } else {
        	throw new FileNotFoundException("File Not Found - " + aFile.getName());
        }
    }

    /**
     * Close the Excel file after reading.
     */
    public void close() {
    	try {
    		if(null != excelPackage) {
    			excelPackage.close();
    		}
		} catch (IOException ioException) {
			ioException.printStackTrace();
		}
    }
    /**
     * @return - Retrieve the first sheet data
     */
    public List<String[]> getRowDataList(){
        return getRowDataList(null, null);
    }

    /**
     * Returns the Workbook data for the requested Row count.
     * @param aRowSize
     * @return - Retrieve the data for the requested RowSize
     */
    public List<String[]> getRowDataList(Integer aRowSize){
        return getRowDataList(null, aRowSize);
    }

    /**
     * Returns List of Map with Key as the Column Header and
     * Value as respective Row value.
     *
     * @return - Retrieve the data with Column Header as key and
     * value with respective row cell value.
     */
    public List<Map<String,String>> getRowHeaderDataMapList(){
        return getRowHeaderDataMapList(null);
    }

    /**
     * Returns List of Map with Key as the Column Header and
     * Value as respective Row value for the requested row count.
     *
     * @param aRowSize
     * @return - Retrieve the data with Column Header as key and
     * value with respective row cell value for the requested row count.
     */
    public List<Map<String,String>> getRowHeaderDataMapList(Integer aRowSize){
    	return getRowHeaderDataMapList(null, aRowSize);
    }
    
    /**
     * Returns List of Map with Key as the Column Header and
     * Value as respective Row value for the requested row count.
     * 
     * @param aSheetName
     * @param aRowSize
     * @return
     */
    public List<Map<String,String>> getRowHeaderDataMapList(String aSheetName, Integer aRowSize){

        List<Map<String,String>> rowHeaderDataMapList = new ArrayList<Map<String,String>>();

        List<String[]> rowDataList = getRowDataList(aSheetName, aRowSize);
        if(!rowDataList.isEmpty()){

            //Assuming the first row is the Header
            String[] headerArr = rowDataList.get(0);

            for(int count = 0; count < rowDataList.size(); count++){

                if(count != 0){

                    Map<String,String> rowDataMap = new LinkedHashMap<>(); //LinkedHashMap to maintain Row order
                    String[] rowDataArr = rowDataList.get(count);
                    
                    int rowDataLength = rowDataArr.length;

                    for(int counter = 0; counter < headerArr.length; counter++){
                    	
                    	String rowData = rowDataLength > counter ? rowDataArr[counter] : null;
                    	
                        rowDataMap.put(headerArr[counter], rowData);
                    }

                    rowHeaderDataMapList.add(rowDataMap);
                }
            }
        }
        
        close();
        
        return rowHeaderDataMapList;

    }

    /**
     * @param aSheetName - Name of the Sheet
     * @param aRowSize - Size of the Row to fetched
     * @return - Retrieve data by SheetName and max rows to fetch
     */
    public List<String[]> getRowDataList(String aSheetName, Integer aRowSize){

        List<String[]> rowDataList = new ArrayList<>();

        List<HashMap<String, List<String[]>>> sheetRowDataMapList = getSheetRowDataMapList(aSheetName, aRowSize, false);
        if(!sheetRowDataMapList.isEmpty()) {

            HashMap<String, List<String[]>> sheetRowDataMap = sheetRowDataMapList.get(0);
            for(String key : sheetRowDataMap.keySet()) {
                rowDataList = sheetRowDataMap.get(key);
                break;
            }
        }

        return rowDataList;
    }

    /**
     * @return - Returns list of HashMap which contains all the row data with sheet name as key
     */
    public List<HashMap<String, List<String[]>>> getAllSheetRowDataMapList(){
        return getSheetRowDataMapList(null, null, true);
    }

    /**
     *
     * @param aSheetName - Name of the Sheet
     * @param aRowSize - Size of the Row to fetched
     * @param fetchAllSheet - Boolean to determine if all sheets should be returned
     * @return - Returns list of HashMap which contains all the row data with sheet name as key
     */
    private List<HashMap<String, List<String[]>>> getSheetRowDataMapList (String aSheetName, Integer aRowSize, boolean fetchAllSheet) {

        List<HashMap<String, List<String[]>>> sheetRowDataMapList = new ArrayList<HashMap<String, List<String[]>>>();

        try {
            XSSFReader.SheetIterator sheetItr = (XSSFReader.SheetIterator) xssfReader.getSheetsData();
            while(sheetItr.hasNext()) {
            	
                InputStream inputStream = sheetItr.next();
                
                String sheetName = sheetItr.getSheetName();
                HashMap<String, List<String[]>> sheetRowDataMap = new HashMap<String, List<String[]>>();

                if(fetchAllSheet) {

                    sheetRowDataMap.put(sheetName, getRowList(inputStream, aRowSize));
                    sheetRowDataMapList.add(sheetRowDataMap);

                } else if(StringUtils.isNotBlank(aSheetName)) {

                    if(sheetName.equals(aSheetName)) {

                        sheetRowDataMap.put(sheetName, getRowList(inputStream, aRowSize));
                        sheetRowDataMapList.add(sheetRowDataMap);
                        break;
                    }

                } else {

                    sheetRowDataMap.put(sheetName, getRowList(inputStream, aRowSize));
                    sheetRowDataMapList.add(sheetRowDataMap);
                    break;
                }

            }

        } catch (Exception exception) {
            System.out.println("Exception Occurred while initializing the Stream Reader");
            exception.printStackTrace();

        }
        return sheetRowDataMapList;
    }

    /**
     * @param aInputStream
     * @param aRowSize
     * @return - Return List of string array for each Row
     * @throws Exception - throws Exception
     */
    public List<String[]> getRowList(InputStream aInputStream, Integer aRowSize) throws Exception{
        List<String[]> rowArrList = new ArrayList<>();
        XMLStreamReader xmlStreamReader = inputFactory.createXMLStreamReader(aInputStream);

        while(xmlStreamReader.hasNext()) {
            xmlStreamReader.next();

            if(xmlStreamReader.isStartElement() && ELEMENT_ROW.equals(xmlStreamReader.getLocalName())) {

                List<String> rowDataList = getRowDataList(xmlStreamReader);
                rowArrList.add(rowDataList.toArray(new String[rowDataList.size()]));

                if(null != aRowSize && rowArrList.size() == aRowSize) {
                    break;
                }
            }
        }
        return rowArrList;
    }

    /**
     * @param aXmlStreamReader - XMLStreamReader input
     * @return - Return List of string cell values
     * @throws XMLStreamException - throws XMLStreamException
     */
    private List<String> getRowDataList(XMLStreamReader aXmlStreamReader) throws XMLStreamException{

        List<String> rowList = new ArrayList<String>();

        while(aXmlStreamReader.hasNext()) {
            aXmlStreamReader.next();

            if(aXmlStreamReader.isStartElement() && ELEMENT_C.equals(aXmlStreamReader.getLocalName())) {
                CellReference cellReference = new CellReference(aXmlStreamReader.getAttributeValue(null, ELEMENT_R));
                if(rowList.size() < cellReference.getCol()) {
                    rowList.add("");
                }

                String cellType = aXmlStreamReader.getAttributeValue(null, ELEMENT_T);
                rowList.add(getCellValue(aXmlStreamReader, cellType));

            } else if(aXmlStreamReader.isEndElement() && ELEMENT_ROW.equals(aXmlStreamReader.getLocalName())) {
                break;
            }
        }
        return rowList;
    }

    /**
     *
     * @param aXmlStreamReader
     * @param aCellType
     * @return - returns the Cell Value
     * @throws XMLStreamException
     */
    private String getCellValue(XMLStreamReader aXmlStreamReader, String aCellType) throws XMLStreamException {
        String cellValue = "";

        while(aXmlStreamReader.hasNext()) {
            aXmlStreamReader.next();

            if(aXmlStreamReader.isStartElement() && ELEMENT_V.equals(aXmlStreamReader.getLocalName())) {

                cellValue = aXmlStreamReader.getElementText();

                if(ELEMENT_S.equals(aCellType)) {
                    int index = Integer.parseInt(cellValue);
                    cellValue = stringsTable.getItemAt(index).toString();
                }

            } else if (aXmlStreamReader.isEndElement() && ELEMENT_C.equals(aXmlStreamReader.getLocalName())) {
                break;
            }
        }

        return cellValue;
    }

}

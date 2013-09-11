//
// JODConverter - Java OpenDocument Converter
// Copyright 2004-2012 Mirko Nasato and contributors
//
// JODConverter is Open Source software, you can redistribute it and/or
// modify it under either (at your option) of the following licenses
//
// 1. The GNU Lesser General Public License v3 (or later)
//    -> http://www.gnu.org/licenses/lgpl-3.0.txt
// 2. The Apache License, Version 2.0
//    -> http://www.apache.org/licenses/LICENSE-2.0.txt
//
package org.artofsolving.jodconverter;

import java.io.File;
import java.util.Enumeration;
import java.util.HashMap;
import java.util.Hashtable;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;
import java.util.logging.Logger;

import org.artofsolving.jodconverter.document.DocumentFamily;
import org.artofsolving.jodconverter.document.DocumentFormat;
import org.artofsolving.jodconverter.office.OfficeException;
import org.artofsolving.jodconverter.util.PlatformUtils;

import com.sun.star.beans.XPropertySet;
import com.sun.star.frame.XModel;
import com.sun.star.lang.XComponent;
import com.sun.star.table.XCell;
import com.sun.star.table.XCellRange;
import com.sun.star.text.XText;
import com.sun.star.text.XTextCursor;
import com.sun.star.text.XTextDocument;
import com.sun.star.text.XTextRange;
import com.sun.star.text.XTextTable;
import com.sun.star.text.XTextTableCursor;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.util.XReplaceDescriptor;
import com.sun.star.util.XReplaceable;
import com.sun.star.util.XSearchDescriptor;
import com.sun.star.util.XSearchable;

public class StandardConversionTask extends AbstractConversionTask {
	
	private final static Logger logger = Logger.getLogger(StandardConversionTask.class .getName()); 
	
    private final DocumentFormat outputFormat;

    private Map<String,?> defaultLoadProperties;
    private DocumentFormat inputFormat;

    public StandardConversionTask(File inputFile, File outputFile,Hashtable parameters,List data, DocumentFormat outputFormat) {
        super(inputFile, outputFile,parameters,data);
        this.outputFormat = outputFormat;
    }

    public void setDefaultLoadProperties(Map<String, ?> defaultLoadProperties) {
        this.defaultLoadProperties = defaultLoadProperties;
    }

    public void setInputFormat(DocumentFormat inputFormat) {
        this.inputFormat = inputFormat;
    }

	@Override
	protected void modifyDocument(XComponent document) throws OfficeException {
		try {
			XTextDocument xTextDocument = (XTextDocument) UnoRuntime.queryInterface(XTextDocument.class, document);
			if (null != this.data && !this.data.isEmpty()) {
				logger.info("################################ start fillTables ####################################");
				fillTables(this.data, xTextDocument);
				logger.info("################################ end fillTables ####################################");
			}
			if (null != this.parameters && !this.parameters.isEmpty()) {
				logger.info("################################ start replaceParameters ####################################");
				replaceParameters(xTextDocument, this.parameters);
				logger.info("################################ end replaceParameters ####################################");
			}
		}
		catch (Exception e) {
			logger.severe(e.getMessage() + " " + e.getCause());
			e.printStackTrace();
			throw new OfficeException(e.getMessage(), e.getCause());
		}
	}

    @Override
    protected Map<String,?> getLoadProperties(File inputFile) {
        Map<String,Object> loadProperties = new HashMap<String,Object>();
        if (defaultLoadProperties != null) {
            loadProperties.putAll(defaultLoadProperties);
        }
        if (inputFormat != null && inputFormat.getLoadProperties() != null) {
            loadProperties.putAll(inputFormat.getLoadProperties());
        }
        return loadProperties;
    }

    @Override
    protected Map<String,?> getStoreProperties(File outputFile, XComponent document) {
        DocumentFamily family = OfficeDocumentUtils.getDocumentFamily(document);
        return outputFormat.getStoreProperties(family);
    }
    
    
    /**
     * @author : dmotta
     * @date : Jul 19, 2013
     * @description : TODO(dmotta) : insert description
     * @return : void TODO(dmotta)
     */
    private void fillTables(List<HashMap> data, XTextDocument xTextDocument) throws Exception {       
          for (HashMap<String, Object> data1 : data) {
            for (Entry<String, Object> entry : data1.entrySet()) {          
              String key = entry.getKey();          
              List list  =  (List)entry.getValue();
              logger.info("key:  " + key + " list:"+list);
              if(null!=key && !key.isEmpty() && null!=list && !list.isEmpty()){
            	  fillData(key, list, xTextDocument);  
              }              
            }
          }
      }

      /**
     * @author : dmotta
     * @date : Jul 19, 2013
     * @description : TODO(dmotta) : insert description
     * @return : void TODO(dmotta)
     */
    private void fillData(String key, List tableData, XTextDocument xTextDocument) throws Exception {

        XSearchable xSearchable = (XSearchable) UnoRuntime.queryInterface(XSearchable.class, xTextDocument);
        XSearchDescriptor xSearchDescriptor = xSearchable.createSearchDescriptor();
        xSearchDescriptor.setSearchString(key);
        Object object = xSearchable.findFirst(xSearchDescriptor);

        XTextRange xTextRange = (XTextRange) UnoRuntime.queryInterface(XTextRange.class, object);
        XTextCursor xTextCursor = xTextRange.getText().createTextCursor();
        XPropertySet xPropertySet = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, xTextCursor);
        Object tableObject = xPropertySet.getPropertyValue("TextTable");
        Object cellObject = xPropertySet.getPropertyValue("Cell");
        XTextTable xTable = (XTextTable) UnoRuntime.queryInterface(XTextTable.class, tableObject);
        XCell xCell = (XCell) UnoRuntime.queryInterface(XCell.class, cellObject);
        XPropertySet xCellPropertySet = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, xCell);
        String cellName = (String) xCellPropertySet.getPropertyValue("CellName");

        String yPositionStr = cellName.substring(PlatformUtils.firstNumericPosition(cellName), cellName.length());
        logger.info("yPositionStr: " + yPositionStr);
        Integer yPosition = Integer.parseInt(yPositionStr);

        xTable.getRows().insertByIndex((yPosition - 1), tableData.size());
        logger.info("cellName: " + cellName);        

        XTextTableCursor xTextTableCursor = xTable.createCursorByCellName(cellName);
        short alt = (short) (tableData.size()-1);
        short anch = (short) (((List)tableData.get(0)).size()-1);
        xTextTableCursor.goRight(anch, true);
        xTextTableCursor.goDown(alt, true);
        logger.info("altura: " + alt + " ancho:" + anch);
        
        String cellsRange = xTextTableCursor.getRangeName();
        logger.info("cellsRange: " + cellsRange);

        XCellRange xCellRange = (XCellRange) UnoRuntime.queryInterface(XCellRange.class, xTable);

        XCellRange newRange = xCellRange.getCellRangeByName(cellsRange);

        for (int i = 0; i < tableData.size(); i++) {
          for (int j = 0; j < ((List)tableData.get(i)).size(); j++) {
            XCell xCellNew = newRange.getCellByPosition(j, i);
            XText xText = (XText) UnoRuntime.queryInterface(XText.class, xCellNew);
            xText.setString((String)((List)tableData.get(i)).get(j));
          }
        }
      }

	/**
	 * @author : dmotta
	 * @date : Jul 19, 2013
	 * @description : TODO(dmotta) : insert description
	 * @return : void TODO(dmotta)
	 */
	private <U extends XModel> void replaceParameters(U u, Hashtable<String, String> parameters) throws Exception {
		XReplaceDescriptor xReplaceDescr = null;
		XReplaceable xReplaceable = null;
		String key = null;
		Enumeration<String> keys = parameters.keys();
		while (keys.hasMoreElements()) {
			key = (String) keys.nextElement();// XTextDocument
			xReplaceable = (XReplaceable) UnoRuntime.queryInterface(XReplaceable.class, u);
			xReplaceDescr = (XReplaceDescriptor) xReplaceable.createReplaceDescriptor();
			xReplaceDescr.setSearchString(key);
			xReplaceDescr.setReplaceString(parameters.get(key));
			xReplaceable.replaceAll(xReplaceDescr);
		}
	}

}

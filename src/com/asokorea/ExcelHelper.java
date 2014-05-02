package com.asokorea;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.io.OutputStreamWriter;
import java.util.ArrayList;
import java.util.Date;
import java.util.Iterator;

import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import javax.xml.xpath.XPath;
import javax.xml.xpath.XPathConstants;
import javax.xml.xpath.XPathExpressionException;
import javax.xml.xpath.XPathFactory;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellReference;
import org.w3c.dom.Document;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;
import org.xml.sax.SAXException;

import sun.misc.Launcher;

public class ExcelHelper {

	private File sourceFile = null;
	private File targetFile = null;
	private Workbook workbook = null;
	private FileOutputStream out = null;
	private OutputStreamWriter osw = null;
	private String xml = null;
	
	public File convertXMLFile(String sourceFilePath) throws InvalidFormatException, IOException{
		
		File file = new File(sourceFilePath);
		
		if(file != null && file.isFile() && file.exists())
		{
			this.sourceFile = file;
		}else{
			String baseDir = Launcher.class.getResource("/").getPath();
			String url = baseDir + sourceFilePath;
			
			file = new File(url);
			
			if(file != null && file.isFile() && file.exists())
			{
				this.sourceFile = file;
			}
		}
		
		this.targetFile = new File(this.sourceFile.getCanonicalPath() + ".xml");
		this.workbook = WorkbookFactory.create(this.sourceFile);
		this.xml = workbook2xml(workbook);

		BufferedWriter bw = null;
		out = new FileOutputStream(this.targetFile);
		osw = new OutputStreamWriter(out, "UTF-8");
			
		bw = new BufferedWriter(osw);
		bw.write(xml);
		bw.flush();
		bw.close();
		
		return this.targetFile;
	}

	public String convertXML(String sourceFilePath) throws InvalidFormatException, IOException{
		
		File file = new File(sourceFilePath);
		
		if(file != null && file.exists() && file.isFile())
		{
			this.sourceFile = file;
		}else{
			String baseDir = Launcher.class.getResource("/").getPath();
			String url = baseDir + sourceFilePath;
			
			file = new File(url);
			
			if(file != null && file.isFile() && file.exists())
			{
				this.sourceFile = file;
			}
		}
		
		this.workbook = WorkbookFactory.create(this.sourceFile);
		this.xml = workbook2xml(workbook);
		return this.xml;
	}
	
	public void close() {
		try {
			this.workbook = null;
			if(this.osw != null){
				this.osw.close();
				this.osw = null;
			}
			if(this.osw != null){
				this.out.close();
				this.out = null;
			}
			this.targetFile = null;
			this.sourceFile = null;
			this.xml = null;
		} catch (IOException e) {
		}
	}
	
	private String cellToString(Cell cell){
		String result = null;

		switch (cell.getCellType()) {
		case Cell.CELL_TYPE_STRING:
			result = cell.getStringCellValue();
			break;
		case Cell.CELL_TYPE_NUMERIC:
			if (DateUtil.isCellDateFormatted(cell)) {
				result = cell.getDateCellValue().toString();
			} else {
				int i = (int)cell.getNumericCellValue();
				double d = cell.getNumericCellValue();

				if(i == d)
				{
					result = String.valueOf(i);
				}else{
					result = String.valueOf(d);
				}
			}
			break;
		case Cell.CELL_TYPE_BOOLEAN:
			result = Boolean.toString(cell.getBooleanCellValue());
			break;
		case Cell.CELL_TYPE_FORMULA:
			result = cell.getCellFormula();
			break;
		default:
			result = null;
		}

		return result;
	}
	
	private String workbook2xml(org.apache.poi.ss.usermodel.Workbook workbook) {
		String result = null;
		StringBuffer sb = null;
		Sheet sheet = null;
		
		if(workbook != null && workbook.getSheetAt(0) != null)
		{
			String newLine = System.getProperty("line.separator");
			
			sb = new StringBuffer();
			sb.append("<?xml version=\"1.0\" ?>");
			sb.append(newLine);
			sb.append("<!DOCTYPE workbook SYSTEM \"workbook.dtd\">");
			sb.append(newLine);
			sb.append(newLine);
			sb.append("<workbook>");
			sb.append(newLine);

			for (int i = 0; i < workbook.getNumberOfSheets(); ++i) {
				
				sheet = workbook.getSheetAt(i);
				
				if(sheet != null && sheet.rowIterator().hasNext()){
					
					sb.append("\t");
					sb.append("<sheet>");
					sb.append(newLine);
					sb.append("\t\t");
					sb.append("<name><![CDATA[" + sheet.getSheetName() + "]]></name>");
					sb.append(newLine);

					int j = 0;
					
					for (Iterator<Row> iterator = sheet.rowIterator(); iterator.hasNext();) {
						Row row = (Row) iterator.next();
						
						int k = 0;

						if(row.getCell(0) != null && row.getCell(0).getStringCellValue() != null
								&& row.getCell(0).getStringCellValue().trim().length() > 0){
							
							sb.append("\t\t");
							sb.append("<row number=\"" + j + "\">");
							sb.append(newLine);

							for (Cell cell : row) {
								sb.append("\t\t\t");
								sb.append("<col number=\"" + k + "\">");
								sb.append("<![CDATA[" + cellToString(cell) + "]]>");
								sb.append("</col>");
								sb.append(newLine);
								k++;
							}
							
							sb.append("\t\t");
							sb.append("</row>");
							sb.append(newLine);
						}

						j++;
						
					}
					
					sb.append("\t");
					sb.append("</sheet>");
					sb.append(newLine);
					
				}
			}
			
			sb.append("</workbook>");
			sb.append(newLine);

			result = sb.toString();
		}
		
		return result;
	}

	public File exportExcel(String sourceFilePath, String targetFilePath, boolean hasTemplete) 
			throws IOException, SAXException, ParserConfigurationException, XPathExpressionException, InvalidFormatException {
		sourceFile = new File(sourceFilePath);
		targetFile = new File(targetFilePath);

		InputStream xmlInputStream = new FileInputStream(sourceFile);
		InputStream excelInputStream = new FileInputStream(targetFile);
		
		Document document = DocumentBuilderFactory.newInstance().newDocumentBuilder().parse(xmlInputStream);
		Workbook workbook = WorkbookFactory.create(excelInputStream);
	
		workbook = bindXml(document, workbook);
		
		document = null;
		xmlInputStream.close();
		xmlInputStream = null;
		excelInputStream.close();
		excelInputStream = null;
		
		OutputStream excelOutputStream = new FileOutputStream(targetFile);
		workbook.write(excelOutputStream);
		excelOutputStream.close();
		excelOutputStream = null;		
		return targetFile;
	}
	
	public Workbook bindXml(Document document, Workbook workbook) throws XPathExpressionException{
		
		XPath xPath = XPathFactory.newInstance().newXPath();
		NodeList cellValueList = (NodeList)xPath.evaluate("//cellValue",document,XPathConstants.NODESET);
		NodeList rowNodeList = (NodeList)xPath.evaluate("//row",document,XPathConstants.NODESET);
		Node rowsNode = (Node)xPath.evaluate("//rows",document,XPathConstants.NODE);
		
		Sheet sheet = workbook.getSheetAt(0);
		
		for (int i = 0; i < cellValueList.getLength(); i++) {
			Node cellValue = cellValueList.item(i);
			String cellName = cellValue.getAttributes().getNamedItem("ref").getTextContent();
			String type = cellValue.getAttributes().getNamedItem("type").getTextContent();
			String value = cellValue.getTextContent();
			CellReference cellRef = new CellReference(cellName);
			Row row = sheet.getRow(cellRef.getRow());
			Cell cell = row.getCell(cellRef.getCol());
			
			if("number".equals(type)){
				double doubleValue = Double.valueOf(value);
				cell.setCellValue(doubleValue);
			}else if("date".equals(type)){
				Date dateValue = new Date(Long.valueOf(value));
				cell.setCellValue(dateValue);
			}else if("bool".equals(type)){
				boolean boolValue = Boolean.valueOf(value);
				cell.setCellValue(boolValue);
			}else if("formula".equals(type)){
				cell.setCellFormula(value);
			}else{
				cell.setCellValue(value);
			}
		}

		if(rowsNode != null && rowNodeList != null && rowNodeList.getLength() > 0)
		{
			CellReference startCellRef = new CellReference(rowsNode.getAttributes().getNamedItem("startRef").getTextContent());
			CellReference endCellRef = new CellReference(rowsNode.getAttributes().getNamedItem("endRef").getTextContent());
			int startRowIndex = startCellRef.getRow();
			int startColIndex = startCellRef.getCol();
			int endColIndex = endCellRef.getCol();
			CellStyle[] cellStyles = new CellStyle[endColIndex+1];
			Row firstRow = sheet.getRow(startRowIndex);
			
			for (int i = startColIndex; i <= endColIndex; i++) {
				cellStyles[i] = firstRow.getCell(i).getCellStyle();
			}
			
			for (int i = startRowIndex; i <= sheet.getLastRowNum(); i++) {
				Row templeteRow = sheet.getRow(i);
				
				if(templeteRow != null){
					sheet.removeRow(templeteRow);					
				}
			}
			
			int rowNodeIndex = 0;
			
			for (int i = startRowIndex; i < startRowIndex + rowNodeList.getLength(); i++) {
				
				Row row = sheet.createRow(i);
				int cellNodeIndex = 0;
				Node rowNode = rowNodeList.item(rowNodeIndex);
				NodeList rowValueNodeList = rowNode.getChildNodes();
				ArrayList<Node> nodes = new ArrayList<Node>();
				
			    for (int idx = 0; idx < rowValueNodeList.getLength(); idx++) {
			        Node currentNode = rowValueNodeList.item(idx);
			        if (currentNode.getNodeType() == Node.ELEMENT_NODE) {
			           nodes.add(currentNode);
			        }
			    }
				
				for (int j = startColIndex; j <= endColIndex; j++) {
					Cell cell = row.createCell(j);
					Node cellNode = nodes.get(cellNodeIndex);
					String type = cellNode.getAttributes().getNamedItem("type").getTextContent();
					String value = cellNode.getTextContent();
					CellStyle cellStyle = cellStyles[j];
					
					cell.setCellStyle(cellStyle);
					
					if("number".equals(type)){
						double doubleValue = Double.valueOf(value);
						cell.setCellValue(doubleValue);
					}else if("date".equals(type)){
						Date dateValue = new Date(Long.valueOf(value));
						cell.setCellValue(dateValue);
					}else if("bool".equals(type)){
						boolean boolValue = Boolean.valueOf(value);
						cell.setCellValue(boolValue);
					}else if("formula".equals(type)){
						cell.setCellType(HSSFCell.CELL_TYPE_FORMULA);
						cell.setCellFormula(value);
					}else if("string".equals(type)){
						if(value != null && value.length() > 0)
						{
							cell.setCellValue(value);
						}else{
							cell.setCellValue("");
						}
					}else{
						cell.setCellValue("");
					}
					
					cellNodeIndex ++;
				}
				rowNodeIndex ++;
			}
		}
		
		return workbook;
	}
}

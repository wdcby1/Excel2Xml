package com.asokorea;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStreamWriter;
import java.util.Iterator;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import sun.misc.Launcher;

public class ExcelHelper {

	private File sourceFile = null;
	private File targetFile = null;
	private Workbook workbook = null;
	private FileOutputStream out = null;
	private OutputStreamWriter osw = null;
	private String xml = null;
	
	public ExcelHelper(String sourceFilePath) throws IOException {
		
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
	}
	
	public File convertXMLFile() throws InvalidFormatException, IOException{
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

	public String convertXML() throws InvalidFormatException, IOException{
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
						
						sb.append("\t\t");
						sb.append("<row number=\"" + j + "\">");
						sb.append(newLine);
						
						int k = 0;
						
						for (Cell cell : row) {
							sb.append("\t\t\t");
							sb.append("<col number=\"" + k + "\">");
							sb.append("<![CDATA[" + cellToString(cell) + "]]>");
							sb.append("</col>");
							sb.append(newLine);
							k++;
						}

						j++;
						
						sb.append("\t\t");
						sb.append("</row>");
						sb.append(newLine);
						
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
}

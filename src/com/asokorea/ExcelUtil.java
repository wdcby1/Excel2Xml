package com.asokorea;

import java.io.IOException;

import javax.xml.parsers.ParserConfigurationException;
import javax.xml.xpath.XPathExpressionException;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.xml.sax.SAXException;

public class ExcelUtil {

	/**
	 * @param argsd
	 */
	public static void main(String[] args) {
		
		String sourceFilePath;
		String targetFilePath;
		ExcelHelper converter = null;
		boolean export = false; 
		String result = null;
		
		if(args != null){
			
			try {
				String mode = args[0];

				if(mode != null && "-export".equals(mode.toLowerCase())){
					// xml to excel
					export = true;
				}else{
					// excel to xml
					export = false;
				}

				sourceFilePath = args[1];
				converter = new ExcelHelper();

				if(export){
					targetFilePath = args[2];
					result = converter.exportExcel(sourceFilePath, targetFilePath, true).getCanonicalPath();
				}else{
					result = converter.convertXML(sourceFilePath);
				}
				
				System.out.print(result);
				System.exit(0);
			} catch (IOException | InvalidFormatException | XPathExpressionException | SAXException | ParserConfigurationException e) {
				e.printStackTrace(System.err);
				System.exit(1);
			} finally {
				if(converter != null){
					converter.close();
				}
			}
		}
	}
}

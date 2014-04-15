package com.asokorea;

import java.io.IOException;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

public class Excel2Xml {

	/**
	 * @param argsd
	 */
	public static void main(String[] args) {
		
		ExcelHelper converter = null;
		boolean saveFile = false; 
		String result = null;
		
		if(args != null){
			
			try {
				String fileName = args[0];
				
				for (String arg : args) {
					if("-f".equals(arg) || "-F".equals(arg)){
						saveFile = true;
						break;
					}
				}
				
				converter = new ExcelHelper(fileName);
				
				if(saveFile){
					result = converter.convertXMLFile().getCanonicalPath();
				}else{
					result = converter.convertXML();
				}
				
				System.out.print(result);
				System.exit(0);
			} catch (IOException | InvalidFormatException e) {
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

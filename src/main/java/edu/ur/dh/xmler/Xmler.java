package edu.ur.dh.xmler;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;
import java.util.LinkedList;
import java.util.List;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerException;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.commons.cli.Options;
import org.apache.commons.cli.CommandLineParser;
import org.apache.commons.cli.CommandLine;
import org.apache.commons.cli.GnuParser;
import org.apache.commons.cli.ParseException;
import org.apache.commons.io.FilenameUtils;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
 

public class Xmler {
	
	public static void main(String[] args) throws IOException, ParserConfigurationException, TransformerException{
		// create Options object
		Options options = new Options();
		
		//file option
		options.addOption("f", true, "Excel 2007 or greater (xlsx file) to process");
		options.addOption("o", true, "output file location");
		options.addOption("n", false, "remove all null or empty string tags");
		options.addOption("h", false, "help");
		
		CommandLineParser parser = new GnuParser();
		CommandLine cmd;
		try {
			cmd = parser.parse( options, args);
			
			if(cmd.hasOption("h")){
				System.out.println("Usage: -f [INPUT_FILE] [ -o [OUTPUT_DIRECTORY] -n ]");
				System.out.println("Where: ");
				System.out.println("-f [INPUTF_FILE] where input file is the name of the 2007 or greater xslx file ");
				System.out.println("-o [OUTPUT_DIRECTORY]  OPTIONAL location where all the sheets converted to xml files will be placed othewise current working directory");
				System.out.println("-n OPTIONAL remove null flag indicates that tags with empty fields will not be created ");

			} else { // missing required options
				if(!cmd.hasOption("f")){
					if(!cmd.hasOption("f")){
						System.out.println("input file option -f is required");
						System.exit(0);
					}
				} else {
					File f = new File(cmd.getOptionValue("f"));
					System.out.println("Looking for input file " + f.getCanonicalPath());
					if(f.exists()){
						File outputDir = null;
						if( cmd.hasOption("o")){ // user has specified an output directory
							outputDir = new File(cmd.getOptionValue("o"));
							if(outputDir.isFile()){ // make sure it is not a file
								System.out.println("The specified OUTPUT directory is a file");
								System.exit(0);
							} 
							
							if(!outputDir.exists()){
								System.out.println("Output directory " + outputDir.getCanonicalPath() + " does not exist");
								System.exit(0);
							}
						} else {
							String workingDir = System.getProperty("user.dir");
							outputDir = new File(workingDir);
							
						}
						System.out.println("Files will be put in " + outputDir.getCanonicalPath());
						Xmler xmler = new Xmler();
						
						Boolean removeEmptyTags = cmd.hasOption("n");
						System.out.println("remove empty tags = " + removeEmptyTags);
						
						xmler.processFile(f, outputDir, removeEmptyTags);
						
						
					}else{
						System.out.println("File " + f.getCanonicalPath() + " not found");
					}
					
				}
			}
			
			
			
			
			
			

		} catch (ParseException e) {
			System.out.println("command line could not be parsed");
		}
		
		//File f = new File("/Users/nathans/may_bragdon_ographies.xlsx");
		
		
	}
	
	/**
	 * 
	 * @param f
	 * @param outputDir
	 * @param removeEmptyTags
	 * @throws IOException
	 * @throws ParserConfigurationException
	 * @throws TransformerException
	 */
	public void processFile(File f, File outputDir, Boolean removeEmptyTags) throws IOException, ParserConfigurationException, TransformerException{
		FileInputStream fis = null;
		XSSFWorkbook myWorkBook = null;
		
		try{
			fis = new FileInputStream(f);
	        // Finds the workbook instance for XLSX file
	        myWorkBook = new XSSFWorkbook (fis);
	       
	        int numSheets = myWorkBook.getNumberOfSheets();
	        for(int x= 0; x < numSheets; x++){
	        	// Return first sheet from the XLSX workbook
		        XSSFSheet mySheet = myWorkBook.getSheetAt(x);
		        processSheet(mySheet, outputDir, removeEmptyTags);
	        }
	        
		} finally{
			fis.close();
			myWorkBook.close();
		}
	    	
	}
	
	/**
	 * Process sheets.
	 * 
	 * @param sheet - sheet to process
	 * @throws ParserConfigurationException 
	 * @throws TransformerException 
	 * @throws IOException 
	 */
	private void processSheet(XSSFSheet sheet, File outputDir, Boolean removeEmptyTags) throws ParserConfigurationException, TransformerException, IOException{
		System.out.println("************** START *******************");
		
		String sheetName = sheet.getSheetName();
		List<String> headers = this.getHeaders(sheet);
		
		
		//content for building xml data
		DocumentBuilderFactory docFactory = DocumentBuilderFactory.newInstance();
		DocumentBuilder docBuilder = docFactory.newDocumentBuilder();
		// root elements
		Document doc = docBuilder.newDocument();
		Element rootElement = doc.createElement(sheetName);
		doc.appendChild(rootElement);
		
		this.processRows(sheet, headers, rootElement, doc, removeEmptyTags);
				
		// write the content into xml file
		TransformerFactory transformerFactory = TransformerFactory.newInstance();
		Transformer transformer = transformerFactory.newTransformer();
		DOMSource source = new DOMSource(doc);
		File outputFile = new File(FilenameUtils.concat(outputDir.getCanonicalPath(), sheetName + ".xml"));
		System.out.println("Writing sheet to: " + outputFile.getCanonicalPath());
		StreamResult result = new StreamResult(outputFile);
 
		transformer.transform(source, result);

		System.out.println("**************** DONE *****************");
	}
	
	/**
	 * Get the headers for the sheet.
	 * This converts all headers to lower case and replaces spaces with underscores
	 * 
	 * @param sheet - sheet to get the header from
	 * @return - list of headers in order
	 */
	private List<String> getHeaders(XSSFSheet sheet){
		XSSFRow row = sheet.getRow(0);
		Iterator<Cell> cellIter = row.cellIterator();
		
		// the first row will hold the columns to process
		List<String> rowData = new LinkedList<String>();
		while(cellIter.hasNext()){
			Cell cell = cellIter.next();
			rowData.add((cell.getStringCellValue().replace(" ", "_").toLowerCase()));
		}
		return rowData;
	}
	
	/**
	 * Process each row with the specified header
	 * 
	 * @param sheet - sheet that contains the rows to process
	 * @param headers - headers extracted from the sheet
	 */
	private void processRows(XSSFSheet sheet, List<String> headers, Element rootElement, Document doc, Boolean removeEmptyTags){
		int numRows = sheet.getPhysicalNumberOfRows();
		System.out.println( "Number of rows = " + numRows);

		for(int rowIndex = 2; rowIndex < numRows; rowIndex++){ // we skip the first row - 
			//this switches to using logical rows (starts at 1)
			// where physical rows starts at zero so we start at two and go from there
			Row aRow = sheet.getRow(rowIndex);
			if( aRow != null ){
				Element rowElement = doc.createElement("row");
				rootElement.appendChild(rowElement);
				
				int dataSize = headers.size();
				for(int cellIndex = 0; cellIndex < dataSize; cellIndex++ ){
					Cell cellData = aRow.getCell(cellIndex);
					
					if( cellData != null ){
						cellData.setCellType(Cell.CELL_TYPE_STRING);
						String cellValue = cellData.getStringCellValue();
						// if keep everything or remove empty tags and tag is not null then add it
						
						if( !removeEmptyTags || (removeEmptyTags &&  !cellValue.trim().equals(""))){
							Element cellElement = doc.createElement(headers.get(cellIndex));
							cellElement.appendChild(doc.createTextNode(cellValue));
							rowElement.appendChild(cellElement);
						} 
					} 
					
				}
			} 
		}
	}
	
	
}

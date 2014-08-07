//Jonathan Pitts - excel to XML transformer 2014
package jon.pitts.excelXML;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;

import org.apache.poi.xssf.usermodel.XSSFSheet; //working with sheets
import org.apache.poi.xssf.usermodel.XSSFWorkbook; //for opening workbooks
import org.apache.poi.ss.usermodel.Row; //excel rows
import org.apache.poi.ss.usermodel.Cell; //excel cells

import org.w3c.dom.Document;
import org.w3c.dom.Element;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerException;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;

//Class excelXML creates an XML file from an xslx file
public class excelXML {

	public static void main(String[] args) {
		//string arraylist of excel headers
		ArrayList<String> headers = new ArrayList<String>();
		//length of arraylist
		int numCol;
		
		try
		{
			//input file
			FileInputStream file = new FileInputStream(new File(args[0]));
			//output file
			String fileOut = args[1];
			
			//Document builder
			DocumentBuilderFactory docFactory = DocumentBuilderFactory.newInstance();
			DocumentBuilder docBuilder = docFactory.newDocumentBuilder();
			
			//root element
			Document doc = docBuilder.newDocument();
			Element rootElement = doc.createElement("Root");
			doc.appendChild(rootElement);
			
			//Create Workbook reference
			XSSFWorkbook workbook = new XSSFWorkbook(file);
			
			//Get first sheet from workbook
			XSSFSheet sheet = workbook.getSheetAt(0);
			
			//iterate through each rows
			Iterator<Row> rowIterator = sheet.iterator(); //row iterator
			
			//get headers from first row
			getHeader(rowIterator, headers);
			numCol = headers.size();
			
			while(rowIterator.hasNext()){
				//create row element
				Element rowElement = doc.createElement("Row");
				rootElement.appendChild(rowElement);
				
				Row row = rowIterator.next();
				
				//iterate through each cell
				//Iterator<Cell> cellIterator = row.cellIterator(); //cell iterator
				//cellIterator does not handle cells with missing policy
				for(int cellIdx = 0; cellIdx < numCol; cellIdx++) {
					Cell cell = row.getCell(cellIdx,Row.CREATE_NULL_AS_BLANK);
					String header = headers.get(cellIdx);
					//header element
					//System.out.println(header);
					
					
					String innerText = "";
					
					//Check cell type and format accordingly
					switch (cell.getCellType()) {
						case Cell.CELL_TYPE_NUMERIC:
							//System.out.print(cell.getNumericCellValue() + "t");
							innerText = String.valueOf(cell.getNumericCellValue());
							break;
						case Cell.CELL_TYPE_STRING:
							//System.out.print(cell.getStringCellValue() + "t");
							innerText = cell.getStringCellValue();
							break;
					}
					//append row to root and skip blank cells
					if(innerText != "") {
						Element headerElement = doc.createElement(header);
						headerElement.appendChild(doc.createTextNode(innerText));
						rowElement.appendChild(headerElement);
					}
				}
			}
			file.close();
			writeXML(doc,fileOut);
		}
		catch(ParserConfigurationException e)
		{
			e.printStackTrace();
		}
		catch(IOException e) {
			e.printStackTrace();
		}
	}
	
	public static void  getHeader(Iterator<Row> rowIterator, ArrayList<String> headers){
		Row row = rowIterator.next();
		
		Iterator<Cell> cellIterator = row.cellIterator(); //cell iterator

		while(cellIterator.hasNext()) {
			Cell cell = cellIterator.next();
			
			//Check cell type
			switch (cell.getCellType()) {
				case Cell.CELL_TYPE_NUMERIC:
					//get cell value and replace spaces with underscore
					String headerNum = String.valueOf(cell.getNumericCellValue()).replaceAll(" ", "_").replaceAll("=","_");
					headers.add(headerNum);
					break;
				case Cell.CELL_TYPE_STRING:
					//get cell value and replace spaces with underscore
					String headerStr = cell.getStringCellValue().replaceAll(" ", "_").replaceAll("=","_");
					headers.add(headerStr);
					break;
			}
		}
	}
	
	public static void writeXML(Document doc,String fileOut) {
		try {
			//Create XML file from document
			TransformerFactory transformerFactory = TransformerFactory.newInstance();
			Transformer transformer = transformerFactory.newTransformer();
			DOMSource source = new DOMSource(doc);
			StreamResult result = new StreamResult(new File(fileOut));
			
			transformer.transform(source, result);
			
			System.out.println("File saved!");
		}
		catch (TransformerException e) {
			e.printStackTrace();
		}
	}

}

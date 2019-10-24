
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.ParseException;


//Apache POI imports
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFHyperlink;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.usermodel.HSSFDataFormatter;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFDataFormat;

//XOM imports
import nu.xom.Document;
import nu.xom.Elements;
import nu.xom.Element;
import nu.xom.Nodes;
import nu.xom.Attribute;
import nu.xom.Serializer;

//imports for later calculations
import java.text.NumberFormat;

public class ExcelXML {


	public static void main(String[] args) {
	

	try {
		
		// Begin iterating through the Excel Spreadsheet as in ExcelReader.java
		// from part 1 of this article, but with a twist:  
		// load the values into XML Elements and track their data types with
		// attributes.
		
		// First, create a new XML Document object to load the Excel sheet into XML.
		// To create an XML Document, first create an Element to be its root.
		
		Element reportRoot = new Element("sheet");
		
		// Create a new XML Document object using the Root element
		Document XMLReport = new Document(reportRoot);
		
		// Set up a FileInputStream to represent the Excel spreadsheet
	
		FileInputStream excelFIS = new FileInputStream("E:\\JEE_DEV\\Employees2\\resources\\Employee_List.xls");
	
		// Create an Excel Workbook Object using the FileInputStream created above 
		HSSFWorkbook excelWB = new HSSFWorkbook(excelFIS);
	
		// Traverse the sheets by looping through sheets, rows, and cells.			
		// Remember, excelWB is the workbook object obtained earlier.
		// Outer Loop:  Loop through each sheet, which would be
		// for (int sheetNumber = 0; sheetNumber < excelWB.getNumberOfSheets(); sheetNumber++) {
		// HSSFSheet oneSheet = excelWB.getSheetAt(sheetNumber);
		// Really, just use the first sheet in the book to keep the example simple.

			HSSFSheet oneSheet = excelWB.getSheetAt(0);

			// Now get the number of rows in the sheet
			int rows = oneSheet.getPhysicalNumberOfRows();				

			// Middle Loop:  Loop through rows in the sheet

			for (int rowNumber = 0; rowNumber < rows; rowNumber++) {
				HSSFRow oneRow = oneSheet.getRow(rowNumber);
				
				// Skip empty (null) rows.
				if (oneRow == null) {
					continue;
				}
				
					
				// Create an XML element to represent the row.
				
				Element rowElement = new Element("row");
				
				// Get the number of cells in the row
				int cells = oneRow.getPhysicalNumberOfCells();

				// Inner Loop:  Loop through each cell in the row

				for (int cellNumber = 0; cellNumber < cells; cellNumber++) {
					HSSFCell oneCell = oneRow.getCell(cellNumber);
		
					// Test the value of the cell.
					// Based on the value type, use the proper 
					// method for working with the value.


			         // If the cell is blank, the cell object is null, so don't 
			         // try to use it.  It will cause errors.
					 // Use continue to skip it and just keep going.
					if (oneCell == null) {
						continue;
					}
					
			
					// Set up a string to use just "header" as the element name
					// to store the column header cells themselves.
					
					String elementName="header";
					
					// Figure out the column position of the cell.
					int cellColumnNumber = oneCell.getColumnIndex();
					
					// If on the first Excel row, don't change the element name from "header".  
					// because the first row is headers.  Before changing the element name,
					// test to make sure you're past the first row.					
					if (rowNumber >0)
					
						// After the first row element stores the column header names
						// in header elements, use those values for element names.  
						// Set the elementName variable equal to the content of
						// the matching column header stored in the first row element.
						// To get the first row element, use getFirstChildElement("row").
				
						// Use the column index number (cellColumnNumber) to select the
						// corresponding header element.  It's a child of the first row,
						// so use getChild(cellColumnNumber) on the first row element. 
						// Get the correct column name from the element using getValue(). 
						// Set the elementName variable equal to the column name.
				
						elementName = reportRoot.getFirstChildElement("row").getChild(cellColumnNumber).getValue();

					// Remove weird characters and spaces from elementName, as they're not allowed in element names.
					elementName = elementName.replaceAll("[\\P{ASCII}]","");
					elementName = elementName.replaceAll(" ", "");

					// Create an XML Element to represent the cell, using 
					// the calculated element Name					
					Element cellElement = new Element(elementName);
					
					// Create an attribute to hold the cell's data format
					// May be repeated for any other formatting item of interest.
					//Attribute dataFormatAttribute = new Attribute("dataFormat", oneCell.getCellStyle().getDataFormatString());

					// Add the Attribute to the cell Element
					//cellElement.addAttribute(dataFormatAttribute);
					
					switch (oneCell.getCellType()) {
					
						case HSSFCell.CELL_TYPE_STRING:
					
							// If the cell value is String, create an attribute
							// for the cellElement to state the data type is a string
							
							//Attribute strTypeAttribute = new Attribute("dataType", "String");					
							//cellElement.addAttribute(strTypeAttribute);
								
							// Append the cell text into the element
							cellElement.appendChild(oneCell.getStringCellValue());
						
							// Append the cell element into the Row
							
							rowElement.appendChild(cellElement);
				
							break;
					
					// Repeat adding the dataType attribute, appending of the cell's data into the element,
					// and appending of the cell into the row for each possible cell data type.
		
						case HSSFCell.CELL_TYPE_FORMULA:
							// If the cell value is Formula, create an attribute
							// for the cellElement to state the data type is a formula
							
							//Attribute formulaTypeAttribute = new Attribute("dataType", "Formula");
						
							// Add the Attribute to the cell
							//cellElement.addAttribute(formulaTypeAttribute);		
							
							// Append the cell data into the element
							cellElement.appendChild(oneCell.getCellFormula());
							
							// Append the cell element into the Row
							rowElement.appendChild(cellElement);
						
							break;

						case HSSFCell.CELL_TYPE_NUMERIC:
							// If the cell value is a number, create an attribute
							// for the cellElement to state the data type is numeric	
							//Attribute cellAttribute = new Attribute("dataType", "Numeric");
						
							// Add the Attribute to the cell
							//cellElement.addAttribute(cellAttribute);
							
							// apply the formatting from the cells to the raw data
							// to get the right format in the XML.  First, create an
							// HSSFDataFormatter object.
							
							HSSFDataFormatter dataFormatter = new HSSFDataFormatter();
							
							// Then use the HSSFDataFormatter to return a formatted string
							// from the cell, rather than a raw numeric value:
							String cellFormatted = dataFormatter.formatCellValue(oneCell);

							//Append the formatted data into the element
							cellElement.appendChild(cellFormatted);
										
							// Append the cell element into the Row
							rowElement.appendChild(cellElement);
							
							break;
							
						case HSSFCell.CELL_TYPE_ERROR:
							// The cell value is Error, create an attribute
							// for the cellElement to state the data type is an error
							
							//Attribute errorTypeAttribute = new Attribute("dataType", "Error");
						
							// Add the Attribute to the cell
							//cellElement.addAttribute(errorTypeAttribute);		
							
							// Append the cell data into the element.
							// It's a byte.  Either convert it, or use a dataFormatter
							cellElement.appendChild(Byte.toString(oneCell.getErrorCellValue()));

							
							// Append the cell element into the Row
							rowElement.appendChild(cellElement);

							break;
					
					}	
					

				// End Inner Loop
				}
				
				// Append the row Element into the Root 
				// if the row isn't empty.  
				if (rowElement.getChildCount()>0) {
					reportRoot.appendChild(rowElement);
				}
			// End Middle Loop	
			}
			
			// Bonus code:  using XPath is fun.  Uncomment the code below
			// for an XPath example of querying based on attributes.
			// For this example, the query is any element with a 
			// dataType attribute equal to "Numeric."
			//
			// Nodes ExampleNodes = reportRoot.query("//*[@dataType='Numeric']");
	        // for (int x = 0; x < ExampleNodes.size(); x++) {
	        //   System.out.println(ExampleNodes.get(x).getValue());
	        // }
			
		// End Outer Sheet Loop, if really looping through sheets
		// }


	
	// To get Employee's salaries, iterate through row elements and get a collection of rows
	
	Elements rowElements = reportRoot.getChildElements("row");

	// For each row element
	
	for (int i = 0; i < rowElements.size(); i++) {
		 
	// Get the salary element, 
	// Calculate 1% of it and store it in a Donation element.
		// Unless it's the first row (0), which needs a header element.
		if (i==0) {
			Element donationElement = new Element("header");
			donationElement.appendChild("Donation");
			
			//Attribute dataFormat = new Attribute("dataFormat","General");
			//donationElement.addAttribute(dataFormat);			
			
			//Attribute dataType = new Attribute("dataType","String");
			//donationElement.addAttribute(dataType);
		
			// Append the donation element to the row element 
			rowElements.get(i).appendChild(donationElement);
		}
		
		// If the row is not the first row, put the donation in the element.
		else {
			Element donationElement = new Element("Donation");

			// The dataFormat attribute of the donation should be 
			// the same number format as salary, which looking at the XML file tells
		    // us is "#,##0".  		  
			//Attribute dataFormat = new Attribute("dataFormat","#,##0");
			//donationElement.addAttribute(dataFormat);

			// Set the dataType attribute of the donation to Numeric.
			//Attribute dataType = new Attribute("dataType","Numeric");
			//donationElement.addAttribute(dataType);
			
			// Get the salary element and its value
			Element salaryElement = rowElements.get(i).getFirstChildElement("Salary");
			String salaryString = salaryElement.getValue();
			
			// Calculate 1% of the salary.  Salary is a string
			// with commas, so it 
			// must be converted for the calculation.
			
			// Get a java.text.NumberFormat object for converting string to a double
			NumberFormat numberFormat = NumberFormat.getInstance(); 
			
			// Use numberFormat.parse() to convert string to double.
			// Throws ParseException.	
	        Number salaryNumber = numberFormat.parse(salaryString);

	        // Use Number.doubleValue() method on salaryNumber to 
	        // return a double to use in the calculation.  And
	        // perform the calculation to figure out 1%.
			double donationAmount = salaryNumber.doubleValue()*.01;
			
			// Append the value of the donation into the donationElement.
			// donationAmount is a double and must be converted to a string.
			donationElement.appendChild(Double.toString(donationAmount));

			// Append the donation element to the row element
			rowElements.get(i).appendChild(donationElement);
			
		//End else
		}

	//End for loop
	}
	
	// Print out the XML version of the spreadsheet to see it in the console
	System.out.println(XMLReport.toXML());
	
	// To save the XML into a file for GEE WHIS, start with a FileOutputStream
	// to represent the file to write, C:\Planet Power\GEE_WHIS.xml.
	FileOutputStream hamsterFile = new FileOutputStream("C:\\Planet Power\\GEE_WHIS.xml");

	// Create a Serializer to handle writing the XML
	Serializer saveTheHamsters = new Serializer(hamsterFile);
	
	// Set child element indent level to 5 spaces to make it pretty
	saveTheHamsters.setIndent(5);
	
	// Write the XML to the file C:\Planet Power\GEE_WHIS.xml
	saveTheHamsters.write(XMLReport);
	
	// Create a new Excel workbook and iterate through the XML 
	// to fill the cells.
	// Create an Excel Workbook Object 
	HSSFWorkbook donationWorkbook = new HSSFWorkbook();

	// Next, create a sheet for the workbook.		
	HSSFSheet donationSheet = donationWorkbook.createSheet(); 
	
	// Iterate through the row elements and then cell elements
	
	// Outer loop:  There was already an Elements collection of all row elements
	// created earlier.  It's called rowElements.  
	// For each row element in rowElements:
	
	for (int j = 0; j < rowElements.size(); j++) {

		// Create a row in the workbook for each row element (j)
		HSSFRow createdRow = donationSheet.createRow(j);
	
		// Get the cell elements from that row element and add them to the workbook.
		Elements cellElements = rowElements.get(j).getChildElements();
	
		// Middle Loop:  Loop through the cell elements.
		for (int k = 0; k < cellElements.size(); k++) {	
	
			// Create cells and cell styles.  Use
			// createCell (int column)
			// The column index is the same as the cell element index, which is k.
			HSSFCell createdCell = createdRow.createCell(k);	

		    // To set the Cell data format, retrieve it from the attribute 
		    // where it was stored: the dataFormat attribute.  Store it in a String.
		 	String dataFormatString = cellElements.get(k).getAttributeValue("dataFormat");
			
			// Create an HSSFCellStyle using the createCellStyle() method of the workbook.	
			HSSFCellStyle currentCellStyle = donationWorkbook.createCellStyle();

			//Create an HSSFDataFormat object from the workbook's method
			HSSFDataFormat currentDataFormat = donationWorkbook.createDataFormat();
			
		    // Get the index of the HSSFDataFormat to use.  The index of the numeric format
		    // matching the dataFormatString is returned by getFormat(dataFormatString).
			short dataFormatIndex = currentDataFormat.getFormat(dataFormatString);
			
		    // Next, use the retrieved index to set the HSSFCellStyle object's DataFormat.
		    currentCellStyle.setDataFormat(dataFormatIndex);	
			
			// Then apply the HSSFCellStyle to the created cell.
			createdCell.setCellStyle(currentCellStyle);
		
			
			// Set cell value and types depending on the dataType attribute
			if (cellElements.get(k).getAttributeValue("dataType")=="String") {
				createdCell.setCellType(HSSFCell.CELL_TYPE_STRING);
				createdCell.setCellValue(cellElements.get(k).getValue());
			}
			
			if (cellElements.get(k).getAttributeValue("dataType")=="Numeric") {
				createdCell.setCellType(HSSFCell.CELL_TYPE_NUMERIC);	
				
				// In this spreadsheet, number styles are times, dates,
				// or salaries.  To store as a number and not as text,
				// salaries should be converted to doubles first.
				// Dates and times should not be converted to doubles first,
				// or you'll be inputting the wrong date or time value.
				// Dates and times can be entered as Java Date objects.
				
				if (cellElements.get(k).getAttributeValue("dataFormat").contains("#")) {
					// If formatting contains a pound sign, it's not a date.
					// Use a Java NumberFormat to format the numeric type cell as a double
					// because like before, the element has commas in it.
					NumberFormat numberFormat = NumberFormat.getInstance(); 
					Number cellValueNumber = numberFormat.parse(cellElements.get(k).getValue());
					createdCell.setCellValue(cellValueNumber.doubleValue());
	
					// Add a hyperlink to the fictional GEE WHIS Web site just
					// to demonstrate that you can.
						HSSFHyperlink hyperlink = new HSSFHyperlink(HSSFHyperlink.LINK_URL);
						hyperlink.setAddress("http://www.ibm.com/developerworks/");
						createdCell.setHyperlink(hyperlink);
				}
				
				else {
					// if it's a date, don't convert to double
					
					createdCell.setCellValue(cellElements.get(k).getValue());
				}
	
			}	
			if (cellElements.get(k).getAttributeValue("dataType")=="Formula") {
				createdCell.setCellType(HSSFCell.CELL_TYPE_FORMULA);
				createdCell.setCellFormula(cellElements.get(k).getValue());
			}	
			
			if (cellElements.get(k).getAttributeValue("dataType")=="Error") {
				createdCell.setCellType(HSSFCell.CELL_TYPE_ERROR);
				// Errors are represented as bytes.	
				createdCell.setCellValue(Byte.parseByte(cellElements.get(k).getValue()));
			}
			
		//End Middle (cell) for loop	
		}	
	// End Outer (row) for loop	
	}
	
	// Demonstrate functions:
	// Add the TODAY() and NOW() functions at bottom of the Excel report
	// to say when the workbook was opened.
	
	//Find the last row and increment by two to skip a row
	int lastRowIndex = donationSheet.getLastRowNum()+2;

	// Create a row and three cells to hold the information.	
	HSSFRow lastRow = donationSheet.createRow(lastRowIndex);
	HSSFCell notationCell = lastRow.createCell(0);
	HSSFCell reportDateCell = lastRow.createCell(1);
	HSSFCell reportTimeCell = lastRow.createCell(2);

	// Set a regular string value in one cell
	notationCell.setCellValue("Time:");
	
	// Setting formula values uses setCellFormula()	
	reportDateCell.setCellFormula("TODAY()");
	reportTimeCell.setCellFormula("NOW()");	

	
	// Create HSSFCellStyle objects for the date and time cells.
	// Use the createCellStyle() method of the workbook.
	
	HSSFCellStyle dateCellStyle = donationWorkbook.createCellStyle();
	HSSFCellStyle timeCellStyle = donationWorkbook.createCellStyle();
	
	// Get a HSSFDataFormat object to set the time and date formats for the cell styles
	HSSFDataFormat dataFormat = donationWorkbook.createDataFormat();
	
	// Set the cell styles to the right format by using the index numbers of
	// the desired formats retrieved from the getFormat() function of the HSSFDataFormat.
	dateCellStyle.setDataFormat(dataFormat.getFormat("m/dd/yy"));
	timeCellStyle.setDataFormat(dataFormat.getFormat("h:mm AM/PM"));

	// Set the date and time cells to the appropriate HSSFCellStyles.
	reportDateCell.setCellStyle(dateCellStyle);
	reportTimeCell.setCellStyle(timeCellStyle);
	
	// Write out the Workbook to a file.  First,
	// you need some sort of OutputStream to represent the file.
	
	FileOutputStream donationOutputStream = new FileOutputStream("C:\\Planet Power\\Employee_Donations.xls");
	
	donationWorkbook.write(donationOutputStream);
	
	// End try
	}
	catch (IOException e) {
		System.out.println("File Input/Output Exception!");
	}
	catch (ParseException e) {
		System.out.println("Number Parse Exception!");
	}
	
	// End main method
	}
	
// End class block
}

import java.io.FileReader;
import java.io.IOException;

// XOM Exception imports
import nu.xom.ParsingException;

// import ValidityException if you'd like to handle it 
// differently from its superclass ParsingException
// import nu.xom.ValidityException;

//XOM imports
import nu.xom.Builder;
import nu.xom.Document;
import nu.xom.Element;



public class XMLReader {

	
	public static void main(String[] args) {
	
	
	try {
		
		FileReader weatherFileReader = new FileReader("C:\\Planet Power\\weather_service.xml");

		// Next, get a nu.xom.Builder object to work with.
		
		Builder weatherBuilder = new nu.xom.Builder();
		
		// Start parsing the file read by FileReader into XML.
		// Use the build() method of nu.xom Builder to read the FileReader 
		// file into a nu.xom.Document object to parse it as XML.
	
		Document weatherDoc = weatherBuilder.build(weatherFileReader);
	
		// To work with the XML Document, get its root element.
	
		Element weatherRoot = weatherDoc.getRootElement();
	
		// Find the dusk and dawn children of the root element.
		Element dawnElement = weatherRoot.getFirstChildElement("dawn");
		Element duskElement = weatherRoot.getFirstChildElement("dusk");
	
		// Print the contents of the dawn and dusk elements
		System.out.println("Hello, sun!  It's "+dawnElement.getValue()+".");
		System.out.println("Good-bye, sun!  It's "+duskElement.getValue()+".");

	// End try
	}
	catch (IOException e) {
		System.out.println("File Input/Output Exception!");
	}
	catch (ParsingException e) {
		System.out.println("XML Parsing Exception!");
	}
	
	// End main method
	}
	
// End class block
}
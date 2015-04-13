package Ict4d;

import jxl.*;
import jxl.read.biff.BiffException;
import jxl.write.*;
import jxl.write.Number;
import jxl.write.biff.RowsExceededException;

import java.io.File;
import java.io.FileWriter;
import java.io.IOException;
import java.io.PrintWriter;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;
import java.util.Locale;

public class Ict4d {

// name: Sander Martijn Kerkdijk
// ICT4D
// VU Amsterdam - Faculty of Exact Science
	
//////////////////////////////////////////////////////////////////////////////////////////////////////////
	
	final static int ROW = 1; 
	final static String BREAKSIZE = "medium";
	
	Ict4d() throws BiffException, IOException, RowsExceededException, WriteException {
		
		            //	Create a workbook object from the file at specified location.
		            //	Change the path of the file as per the location on your computer.
					//	Calls the VXML Class
		Workbook wrk1 =  Workbook.getWorkbook(new File("/Users/Sander/Desktop/market.xls"));
		Sheet sheet1 = wrk1.getSheet(0);
		makeVXML(sheet1);      	
	}
	
	static String MenuChoicesMarkets(Sheet sheet1,int column) {
					//	Creates the choices menu for markets
			        //	
					//
					//	<Menu Id="{Id of menu}" dtmf={false/true}>
					//	<prompt>
					// 	{Text that need to be spoken out for introduction of menu}
					//	<break size = "{short/medium/long}"/>
					//	Fetch in a FOR loop all markets {
					//	{number of market}.{name of market}
					// 	<break size = "{short/medium/long}"/>
					//	}
					//	</prompt>
					//	Fetch in a For loop all declarations of choices {
					//	<choice next ="#{name of market}"/>
					// }
					// </menu>
					//
					//
		Cell[] cell = sheet1.getColumn(column);
		String Choices ="<prompt>\n";
		
		Choices +="Select one of the following markets\n ";
		Choices +="<break size=\""+BREAKSIZE+"\"/>\n";
		for(int i = column ; i < (sheet1.getColumn(column).length);i++) {	
			Choices += i+"."+cell[i].getContents() +"\n";
			Choices +="<break size=\"medium\"/>\n";
		}
		Choices +="</prompt>\n";
		for(int i = column ; i < (sheet1.getColumn(column).length);i++) {	
			Choices += "<choice next=\"#"+ cell[i].getContents()+"\"/>\n";
		}
		Choices +="</menu>\n\n";
		return Choices;
	}

	static String menuChoicesProducts(Sheet sheet1,int Row,int places) {
					// 	Creates the choices menu for products
					//	
					// 	<Menu Id="{Id of menu}" dtmf={false/true}>
					// 	<prompt>
					//	{Text that need to be spoken out for introduction of menu}
					//	<break size = "{short/medium/long}"/>
					//	Fetch in a FOR loop all products {
					//	{{number of product}.{name of product}}
					//	<break size = "{short/medium/long}"/>
					// 	{text for knowing all the prices}
					//	}
					//	</prompt>
					//	Fetch in a FOR loop all declarations of choices {
					//	<choice next ="#{{name of market}_{name of product}}"/>
					// }
					// <choice next ="#{{name of market}_all"/>
					// </menu>
					//
					// 

		String MenuChoicesProducts = "";
		Cell[] cell = sheet1.getRow(0);
		Cell[] cell2 = sheet1.getColumn(Row);
		MenuChoicesProducts +=	"<menu id=\""+cell2[places].getContents()+ "\" dtmf=\"true\">\n";
		MenuChoicesProducts += "<prompt>\n"
		+ "Select between\n"
		+ "<break size=\""+BREAKSIZE+"\"/>\n";
		for(int i = Row+1 ; i < (sheet1.getRow(Row).length);i++) {
			MenuChoicesProducts +=  (i-1)+"."+cell[i].getContents() +"\n";	
			MenuChoicesProducts +="<break size=\""+BREAKSIZE+"\"/>\n";
		}
		MenuChoicesProducts += "Or Select "+(sheet1.getRow(Row).length-1)+" to know all the prices";
		MenuChoicesProducts += "</prompt>";
		for(int i = Row+1 ; i < (sheet1.getRow(Row).length);i++) {
			MenuChoicesProducts += "\n<choice next=\"#"+cell2[places].getContents()+ "_" +cell[i].getContents()+"\"/>";	
		}
		MenuChoicesProducts += "\n<choice next=\"#"+cell2[places].getContents()+ "_All\"/>";
		MenuChoicesProducts +="\n</menu>\n";
		return MenuChoicesProducts;
	}

	
	static String Date () {
					//
					//
					// construct the date String
					//{{name of day},{day number} {month} {year} {hour} {minute} }
					//
					//
		
		Date now = new Date(); 
		String Date = new SimpleDateFormat("EEEE, d MMMM yyyy HH:mm",Locale.ENGLISH).format(now);
		
		return Date;
	}
	
	static String getPricesMarkets(Sheet sheet1,int Row, int places) {
					//	 Creates the form for market prices
					//	
					//	Fetch in a FOR loop all the market prices in a form structure {
					//	
					// 	<form id ="{{name of market_{name of product / or All}}">
					// 	<block>
					// 	<prompt>
					//	{The price of {product} for {{name of day},{day number} {month} {year} {hour} {minute} } is  {price} CEDI
					//	<break size = "{short/medium/long}"/>
					//	</prompt>
					//	<goto next = "#{name of main menu}"/>
					//	</block>
					//	</form>
					// }
					// 
		Cell[] cell = sheet1.getRow(0);
		Cell[] cell2 = sheet1.getColumn(Row);
		Cell[] cell3 = sheet1.getRow(places);
		String date = Date();
		String getPricesMarkets = "";
		String getPricesAllMarkets="\n<form id=\""+cell2[places].getContents()+ "_All\"> \n" 
		+" <block>\n"
		+"<prompt>\n"
		+"The Prices for "+cell2[places].getContents()+" at " +date+ " are \n"
		+ "<break size=\""+BREAKSIZE+"\"/>\n";
		
		String End = "</prompt>\n"	
		+"<goto next=\"#menu1\"/>\n"	
		+"</block>\n" 
		+ "</form>\n";
		
		
		for(int i = Row+1 ; i < (sheet1.getRow(Row).length);i++) {
			String Begin = "\n<form id=\""+cell2[places].getContents()+ "_" +cell[i].getContents()+"\"> \n" 
			+" <block>\n"
			+"<prompt>\n"; 
			getPricesMarkets += Begin; 
			getPricesMarkets +="The price of "+cell[i].getContents()+" for "+date+" is "+ cell3[i].getContents() +" CEDI \n"
			+ "<break size=\""+BREAKSIZE+"\"/>\n";
			getPricesAllMarkets += cell[i].getContents()+" is " + cell3[i].getContents() +" CEDI \n"
			+ "<break size=\""+BREAKSIZE+"\"/>\n";
			getPricesMarkets+= End;		
		}
		getPricesAllMarkets +=End;
		getPricesMarkets +=getPricesAllMarkets;
		
		return getPricesMarkets;
	}

	static void makeVXML(Sheet sheet1) throws IOException, RowsExceededException, WriteException {
		
					//	 Creates VXML
					//	
					//	Makes new output object ("/Users/Sander/Desktop/output_voice.xml)
					//	
					// 	<?xml version="1.0" encoding="UTF-8"?>
					//	<vxml version = "2.1" >
					//  Fetch MenuChoicesMarkets
					//
					//	Fetch in a FOR loop all the menuChoicesProducts and getPricesMarkets and put a newline after it.
					//
					// 	Read out to output_voice.xml
					//	Close document
					//
					//
		
		PrintWriter out = new PrintWriter(new FileWriter("/Users/Sander/Desktop/output_voice.xml", false), true);
		String VXML ="";
		String Begin = "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n"
		+ "<vxml version = \"2.1\" >\n\n\n"
		+ "<menu id=\"menu1\" dtmf=\"true\">\n";
		String End = "\n </vxml>"; 
		VXML +=Begin;
		VXML +=MenuChoicesMarkets(sheet1,ROW);
		for(int i = 1 ; i < (sheet1.getColumn(ROW).length);i++){	
			VXML +=	menuChoicesProducts(sheet1,ROW,i);
			VXML += getPricesMarkets(sheet1,ROW,i);
			VXML +="\n";
		}
		VXML += End;
		out.write(VXML);
		out.close();	 
	}
	void start() {	
	}
	public static void main(String[] args) throws BiffException, IOException, RowsExceededException, WriteException {
		new Ict4d().start();
	}		
}

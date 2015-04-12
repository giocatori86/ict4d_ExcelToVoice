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
import java.util.Calendar;

public class Ict4d {

// name: Sander Martijn Kerkdijk
// ICT4D
// VU Amsterdam - Faculty of Exact Science
	
//////////////////////////////////////////////////////////////////////////////////////////////////////////
	
public static String Choices = "";
public static String Form ="";
public static String Prices ="";
	
	
Ict4d() throws BiffException, IOException, RowsExceededException, WriteException {
	
		            //Create a workbook object from the file at specified location.
		            //Change the path of the file as per the location on your computer.
		            Workbook wrk1 =  Workbook.getWorkbook(new File("/Users/Sander/Desktop/market.xls"));
		            Sheet sheet1 = wrk1.getSheet(0);
		            print(sheet1);      
		            
		             
}

static String MenuChoicesMarkets(Sheet sheet1,int column) {
	Cell[] cell = sheet1.getColumn(column);
	for(int i = column ; i < (sheet1.getColumn(column).length);i++) {	
	Choices += "<choice next=\"#"+ cell[i].getContents()+"\"/>\n";
	}
	Choices +="</menu>\n\n";
	return Choices;
}

static String MenuChoicesProducts(Sheet sheet1,int Row,int places) {
	String MenuChoicesProducts = "";
	Cell[] cell = sheet1.getRow(0);
	Cell[] cell2 = sheet1.getColumn(Row);
	MenuChoicesProducts +=	"<menu id=\"products_"+cell2[places].getContents()+ "\" dtmf=\"true\">\n";
	for(int i = Row+1 ; i < (sheet1.getRow(Row).length);i++) {
		MenuChoicesProducts += "<choice next=\"#"+cell2[places].getContents()+ "_" +cell[i].getContents()+"\"/>\n";	
	}
	MenuChoicesProducts +="</menu>\n";
	return MenuChoicesProducts;
}


static String getPricesMarkets(Sheet sheet1,int Row, int places) {
	String getPricesMarkets = "";
	Cell[] cell = sheet1.getRow(0);
	Cell[] cell2 = sheet1.getColumn(Row);
	Cell[] cell3 = sheet1.getRow(places);
	
	String End = "</prompt>\n"
	 +"</block>\n" 
	+ "</form>\n";
	for(int i = Row+1 ; i < (sheet1.getRow(Row).length);i++) {
		String Begin = "\n<form id=\""+cell2[places].getContents()+ "_" +cell[i].getContents()+"\"> \n" 
				 +" <block>\n"
				+"<prompt>\n"; 
		getPricesMarkets += Begin;
		getPricesMarkets +="The price of "+cell[i].getContents()+" is today "+ cell3[i].getContents() +" schepjes";
		getPricesMarkets+= End;		
	}

	return getPricesMarkets;
}


static String FormBuilderPrices(Sheet sheet1,int row) {
	String FormBuilderPrices = "";
	Cell [] cell = sheet1.getRow(row);
	for(int i = 2 ; i < (sheet1.getRow(row).length);i++){
		FormBuilderPrices += "<form id=\"" + cell[i].getContents()+"\">\n";
		FormBuilderPrices += "<block> \n";
		FormBuilderPrices +="</block> \n</form>\n\n";
}
	return FormBuilderPrices;
}

static String FormBuilderMarkets(Sheet sheet1,int row) {
	String FormBuilderMarket = "";
	Cell [] cell = sheet1.getColumn(row);
	for(int i = 1 ; i < (sheet1.getColumn(row).length);i++){	
	FormBuilderMarket += "<form id=\"" + cell[i].getContents()+"\">\n";
	FormBuilderMarket += "<block> \n";
	FormBuilderMarket +=MenuChoicesProducts(sheet1,row,i);
	FormBuilderMarket +="</block> \n</form>\n\n";
	FormBuilderMarket += getPricesMarkets(sheet1,row,i);
}
	return FormBuilderMarket;
}

static void print(Sheet sheet1) throws IOException, RowsExceededException, WriteException {
	 PrintWriter out = new PrintWriter(new FileWriter("output.vxml", true), true);
	 String VXML ="";
	 String Begin = "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n"
			 + "<vxml version = \"2.1\" >\n "
	         + "<menu id=\"menu1\" dtmf=\"true\">\n"
	         + "<prompt>Give in your given code </prompt>\n";
	 String End = "\n </vxml>"; 
	VXML +=Begin;
	VXML +=MenuChoicesMarkets(sheet1,1);
	VXML +=FormBuilderMarkets(sheet1,1);
	VXML += End;
	
	
	System.out.print(VXML);
	 

	 
}
		 


	
		      
		
		
void start() {
	 
}
	public static void main(String[] args) throws BiffException, IOException, RowsExceededException, WriteException {
		new Ict4d().start();
	}	
	
	
}

package com.WordExtractor;

import java.io.FileInputStream;
import java.io.IOException;
import java.sql.SQLException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;


public class WordDataExtract {

	public void getWordData(XWPFDocument xdoc) throws SQLException {
		List<XWPFTable>  tables = xdoc.getTables();
		tables = xdoc.getTables();	

		ArrayList<String> data = new ArrayList<String>();		

		for (int t=0; t<=tables.size()-1; t++) {

			XWPFTable table = tables.get(t);			

			for(int r=0; r<=table.getRows().size()-1;r++ )
			{
				XWPFTableRow row = table.getRows().get(r);					

				for(int c=0; c<= row.getTableCells().size()-1; c++) {

					XWPFTableCell cell = row.getTableCells().get(c);			

					StringBuffer cell1 = new StringBuffer();
					String service = null,description = null;


					if(c==0 && (!cell.getText().equalsIgnoreCase("") && cell.getText() != null)) {
						service = "Service= "+cell.getText();				
						data.add(service);						

					}if(c==1 && !cell.getText().equalsIgnoreCase("")) {
						cell1= cell1.append(cell.getText());		
						description = "Description= "+cell1.toString().trim();
						data.add(description);						
					}
				}				
			}
		}

		Iterator<String> iterator = data.iterator();		
		ArrayList<Data> newdata = new ArrayList<>();
		Data serviceObj = null;

		while (iterator.hasNext()) {
			String val = iterator.next();

			if(val.startsWith("Service=")) {				
				if(serviceObj != null) {
					newdata.add(serviceObj);
				}
				serviceObj = new Data();
				String serviceName = val.substring(8);
				serviceObj.setServiceName(serviceName);
			}
			if(val.startsWith("Description=")) {		
				serviceObj.getDescription().append(val.substring(12));
			}			
		}

		newdata.add(serviceObj);
		
		for(int i=0;i<=newdata.size()-1;i++) {
			System.out.println(newdata.get(i).getServiceName());
			System.out.println(newdata.get(i).getDescription());
			System.out.println("**********");
		}	

	}

	public static void main(String[] args) throws InvalidFormatException, IOException, SQLException
	{
		FileInputStream fis = new FileInputStream("C://Inputfile.docx");

		XWPFDocument xdoc = new XWPFDocument(OPCPackage.open(fis));
		WordDataExtract obj = new WordDataExtract();
		obj.getWordData(xdoc);

		xdoc.close();

	}

}

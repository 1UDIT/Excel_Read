/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Main.java to edit this template
 */
package excelread;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStreamReader;
import java.io.OutputStreamWriter;
import java.io.StringReader;
import java.net.HttpURLConnection;
import java.net.URL;
import java.util.Iterator;
import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.w3c.dom.NodeList;
import org.w3c.dom.Document;
import org.xml.sax.InputSource;

/**
 *
 * @author udit
 */
public class ExcelRead {

    /**
     * @param args the command line arguments
     * @throws java.io.IOException
     */
    public static void main(String args[]) throws IOException {
        try {
//            File file = new File(args[0]);
            FileInputStream fis = new FileInputStream(new File("C:\\Users\\udits\\OneDrive\\Desktop\\ReadFile\\ColorsKannada+Super_May2022.xlsx"));
//            FileInputStream fis = new FileInputStream(new File(args[0]));

            XSSFWorkbook workbook = new XSSFWorkbook(fis);
            XSSFSheet spreadsheet = workbook.getSheetAt(0);

            Iterator< Row> rowIterator = spreadsheet.iterator();
            Row headerRow = spreadsheet.getRow(0); // first row contains column names
            String columnName = "ID"; // the name of the column you want to print

            int columnIndex = -1;
            for (Cell cell : headerRow) {
                if (cell.getStringCellValue().equalsIgnoreCase(columnName)) {
                    columnIndex = cell.getColumnIndex();
                    break;
                }
            }

            if (columnIndex == -1) {
                System.out.println("Column not found");
                return;
            }

            for (Row row : spreadsheet) {
                Cell cell = row.getCell(columnIndex);

                if (cell != null && cell.getRowIndex() >= 1) {
//                    System.out.println(cell.getStringCellValue());

//                  String url = "http://192.168.1.100:9763/services/DIVArchiveWS_SOAP_2.1.DIVArchiveWSHttpSoap12Endpoint/";
                    String url = "http://192.168.1.101:9763/services/DIVArchiveWS_SOAP_2.1.DIVArchiveWSHttpSoap12Endpoint/";
                    URL obj = new URL(url);
                    HttpURLConnection con = (HttpURLConnection) obj.openConnection();
                    con.setRequestMethod("POST");
                    con.setRequestProperty("Content-Type", "application/xml");
//            con.setRequestProperty("accept-encoding","none");
                    String xml = "<ns1:getObjectDetailsList xmlns:ns1=\"http://interaction.api.ws.diva.fpdigital.com/xsd\">\n"
                            + "    <ns1:sessionCode>test</ns1:sessionCode>\n"
                            + "    <ns1:isFirstTime>1</ns1:isFirstTime>\n"
                            + "    <ns1:initialTime>0</ns1:initialTime>\n"
                            + "    <ns1:listType >1</ns1:listType>\n"
                            + "    <ns1:objectsListType >2</ns1:objectsListType>\n"
                            + "    <ns1:listPosition >0</ns1:listPosition>\n"
                            + "    <ns1:maxListSize>50</ns1:maxListSize>\n"
                            + "    <ns1:objectName>" + cell.getStringCellValue() + "</ns1:objectName>\n"
                            + "    <ns1:objectCategory>*</ns1:objectCategory>\n"
                            + "    <ns1:mediaName>*</ns1:mediaName>\n"
                            + "    <ns1:levelOfDetail>0</ns1:levelOfDetail>\n"
                            + "</ns1:getObjectDetailsList>";
                    con.setDoOutput(true);
//            System.out.println(xml);
                    try ( OutputStreamWriter wr = new OutputStreamWriter(con.getOutputStream())) {
                        wr.write(xml);
                        wr.flush();
                    } catch (IOException e) {
                        System.out.println(e);
                        e.printStackTrace();
                    }
//                    String responseStatus = con.getResponseMessage();
//                    System.out.println("Response :" + responseStatus);
                    BufferedReader in = new BufferedReader(new InputStreamReader(con.getInputStream()));
                    String responseLine;
                    StringBuilder response = new StringBuilder();
                    StringBuilder responseNull = new StringBuilder();
                    while ((responseLine = in.readLine()) != null) {
                        DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
                        DocumentBuilder builder;
                        InputSource is;
                        try {
                            builder = factory.newDocumentBuilder();
                            is = new InputSource(new StringReader(responseLine));
                            Document doc = builder.parse(is);
                            NodeList list = doc.getElementsByTagName("objectName");
//                            responseNull.append(list.item(0).getTextContent());
                            if (list.item(0) != null) {
                                response.append(responseLine);
                                System.out.println(cell.getStringCellValue());
                                System.out.println("response:" + response);
                                System.out.println("\n");
                            }else{
                                System.out.println("Response not exist"+" "+cell.getStringCellValue());
                                System.out.println("\n");
                            }
                        } catch (ParserConfigurationException e) {
                            System.out.println("e" + e);
                        }
                    }
                    in.close();

                }
            }

//            StringBuilder sb = new StringBuilder();
//            for (Row row : spreadsheet) { // For each Row.
//                Cell cell = row.getCell(1); // Get the Cell at the Index / Column you want.
//                sb.append(row.getCell(1));
//                sb.append("\n");
//            }
//
//            System.out.println(sb.toString());
//            while (rowIterator.hasNext()) {
//                XSSFRow row = (XSSFRow) rowIterator.next();
//                  StringBuilder sb = new StringBuilder();
//                Iterator< Cell> cellIterator = row.cellIterator();
//                while (cellIterator.hasNext()) {
//                    Cell cell = cellIterator.next();
//
//                    switch (cell.getCellType()) {
//                        case Cell.CELL_TYPE_NUMERIC:
//                            sb.append(Integer.toString(Double.valueOf(cell.getNumericCellValue()).intValue()));
//                            System.out.print(Integer.toString(Double.valueOf(cell.getNumericCellValue()).intValue()) + " \t\t ");
//                            break;
//
//                        case Cell.CELL_TYPE_STRING:
////                             sb.append(cell.getStringCellValue());
//                            System.out.print(cell.getStringCellValue() + " \t\t ");
//                            break;
//                    }
//                }
//                System.out.println(); 
//                System.out.println(sb.toString());
//            }
//            fis.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}



//System.out.print(Integer.toString(Double.valueOf(cell.getNumericCellValue()).intValue()) + "\t\t\t");

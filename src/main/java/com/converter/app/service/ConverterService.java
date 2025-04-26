package com.converter.app.service;

import java.io.ByteArrayOutputStream;
import java.util.List;

import org.docx4j.convert.in.xhtml.XHTMLImporterImpl;
import org.docx4j.jaxb.Context;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.StyleDefinitionsPart;
import org.springframework.stereotype.Service;

@Service
public class ConverterService {

	String html = "<!DOCTYPE html>\r\n"
	        + "<html lang=\"en\">\r\n"
	        + "<head>\r\n"
	        + "  <meta charset=\"UTF-8\" />\r\n"
	        + "  <meta name=\"viewport\" content=\"width=device-width, initial-scale=1.0\" />\r\n"
	        + "  <title>Report</title>\r\n"
	        + "  <style>\r\n"
	        + "    body {\r\n"
	        + "      font-family: Arial, sans-serif;\r\n"
	        + "      background-color: #f4f7fa;\r\n"
	        + "      margin: 0;\r\n"
	        + "      padding: 20px;\r\n"
	        + "    }\r\n"
	        + "    h1, h2 {\r\n"
	        + "      color: #2c3e50;\r\n"
	        + "    }\r\n"
	        + "    .container {\r\n"
	        + "      width: 80%;\r\n"
	        + "      margin: 0 auto;\r\n"
	        + "    }\r\n"
	        + "    table {\r\n"
	        + "      width: 100%;\r\n"
	        + "      margin-bottom: 20px;\r\n"
	        + "      border-collapse: collapse;\r\n"
	        + "      background-color: white;\r\n"
	        + "    }\r\n"
	        + "    table, th, td {\r\n"
	        + "      border: 1px solid #ddd;\r\n"
	        + "    }\r\n"
	        + "    th, td {\r\n"
	        + "      padding: 10px;\r\n"
	        + "      text-align: left;\r\n"
	        + "    }\r\n"
	        + "    th {\r\n"
	        + "      background-color: #34495e;\r\n"
	        + "      color: white;\r\n"
	        + "    }\r\n"
	        + "    td {\r\n"
	        + "      background-color: #f9f9f9;\r\n"
	        + "    }\r\n"
	        + "    caption {\r\n"
	        + "      font-size: 1.5em;\r\n"
	        + "      margin-bottom: 10px;\r\n"
	        + "      color: #2980b9;\r\n"
	        + "    }\r\n"
	        + "    .section {\r\n"
	        + "      margin-bottom: 40px;\r\n"
	        + "    }\r\n"
	        + "    .section h2 {\r\n"
	        + "      font-size: 1.8em;\r\n"
	        + "      color: #2980b9;\r\n"
	        + "      border-bottom: 2px solid #2980b9;\r\n"
	        + "      padding-bottom: 5px;\r\n"
	        + "    }\r\n"
	        + "    .footer {\r\n"
	        + "      text-align: center;\r\n"
	        + "      font-size: 0.9em;\r\n"
	        + "      color: #7f8c8d;\r\n"
	        + "      margin-top: 40px;\r\n"
	        + "    }\r\n"
	        + "  </style>\r\n"
	        + "</head>\r\n"
	        + "<body>\r\n"
	        + "  <div class=\"container\">\r\n"
	        + "    <h1>Report Title</h1>\r\n"
	        + "    <div class=\"section\">\r\n"
	        + "      <h2>Section 1: Overview</h2>\r\n"
	        + "      <p>This section provides an overview of the data.</p>\r\n"
	        + "      <table>\r\n"
	        + "        <caption>Summary of Key Metrics</caption>\r\n"
	        + "        <thead>\r\n"
	        + "          <tr>\r\n"
	        + "            <th>Metric</th>\r\n"
	        + "            <th>Value</th>\r\n"
	        + "            <th>Unit</th>\r\n"
	        + "          </tr>\r\n"
	        + "        </thead>\r\n"
	        + "        <tbody>\r\n"
	        + "          <tr>\r\n"
	        + "            <td>Revenue</td>\r\n"
	        + "            <td>1,000,000</td>\r\n"
	        + "            <td>USD</td>\r\n"
	        + "          </tr>\r\n"
	        + "          <tr>\r\n"
	        + "            <td>Cost</td>\r\n"
	        + "            <td>800,000</td>\r\n"
	        + "            <td>USD</td>\r\n"
	        + "          </tr>\r\n"
	        + "          <tr>\r\n"
	        + "            <td>Profit</td>\r\n"
	        + "            <td>200,000</td>\r\n"
	        + "            <td>USD</td>\r\n"
	        + "          </tr>\r\n"
	        + "        </tbody>\r\n"
	        + "      </table>\r\n"
	        + "    </div>\r\n"
	        + "    <div class=\"section\">\r\n"
	        + "      <h2>Section 2: Detailed Analysis</h2>\r\n"
	        + "      <p>This section delves into the specifics of the data points.</p>\r\n"
	        + "      <table>\r\n"
	        + "        <caption>Sales Data by Region</caption>\r\n"
	        + "        <thead>\r\n"
	        + "          <tr>\r\n"
	        + "            <th>Region</th>\r\n"
	        + "            <th>Sales</th>\r\n"
	        + "            <th>Target</th>\r\n"
	        + "            <th>% Achievement</th>\r\n"
	        + "          </tr>\r\n"
	        + "        </thead>\r\n"
	        + "        <tbody>\r\n"
	        + "          <tr>\r\n"
	        + "            <td>North</td>\r\n"
	        + "            <td>500,000</td>\r\n"
	        + "            <td>550,000</td>\r\n"
	        + "            <td>91%</td>\r\n"
	        + "          </tr>\r\n"
	        + "          <tr>\r\n"
	        + "            <td>South</td>\r\n"
	        + "            <td>400,000</td>\r\n"
	        + "            <td>400,000</td>\r\n"
	        + "            <td>100%</td>\r\n"
	        + "          </tr>\r\n"
	        + "          <tr>\r\n"
	        + "            <td>East</td>\r\n"
	        + "            <td>300,000</td>\r\n"
	        + "            <td>350,000</td>\r\n"
	        + "            <td>86%</td>\r\n"
	        + "          </tr>\r\n"
	        + "          <tr>\r\n"
	        + "            <td>West</td>\r\n"
	        + "            <td>200,000</td>\r\n"
	        + "            <td>250,000</td>\r\n"
	        + "            <td>80%</td>\r\n"
	        + "          </tr>\r\n"
	        + "        </tbody>\r\n"
	        + "      </table>\r\n"
	        + "    </div>\r\n"
	        + "    <div class=\"footer\">\r\n"
	        + "      <p>Report generated on: April 26, 2025</p>\r\n"
	        + "    </div>\r\n"
	        + "  </div>\r\n"
	        + "</body>\r\n"
	        + "</html>\r\n";


	
	public byte[] convertToDocx() throws Exception {

	    // Create the Word document package
	    WordprocessingMLPackage mlPackage = WordprocessingMLPackage.createPackage();

	    // Ensure the document has a StyleDefinitionsPart
	    if (mlPackage.getMainDocumentPart().getStyleDefinitionsPart() == null) {
	        StyleDefinitionsPart styleDefinitionsPart = new StyleDefinitionsPart();
	        styleDefinitionsPart.setJaxbElement(
	            Context.getWmlObjectFactory().createStyles()
	        );
	        mlPackage.getMainDocumentPart().addTargetPart(styleDefinitionsPart);
	    }

	    // Import the HTML into the Word document
	    XHTMLImporterImpl importer = new XHTMLImporterImpl(mlPackage);
	    List<Object> elements = importer.convert(html, null);
	    mlPackage.getMainDocumentPart().getContent().addAll(elements);

	    // Write the document to a byte array and return it
	    try (ByteArrayOutputStream arrayOutputStream = new ByteArrayOutputStream()) {
	        mlPackage.save(arrayOutputStream);
	        return arrayOutputStream.toByteArray();
	    } catch (Exception e) {
	        throw e;
	    }
	}

	
}

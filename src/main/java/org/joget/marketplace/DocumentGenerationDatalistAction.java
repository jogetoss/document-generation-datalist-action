/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Class.java to edit this template
 */
package org.joget.marketplace;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.net.URLEncoder;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Map;
import java.util.Set;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;
import javax.servlet.ServletException;
import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.extractor.XWPFWordExtractor;
import org.apache.poi.xwpf.usermodel.Document;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.joget.apps.app.model.AppDefinition;
import org.joget.apps.app.service.AppPluginUtil;
import org.joget.apps.app.service.AppResourceUtil;
import org.joget.apps.app.service.AppService;
import org.joget.apps.app.service.AppUtil;
import org.joget.apps.datalist.model.DataList;
import org.joget.apps.datalist.model.DataListActionDefault;
import org.joget.apps.datalist.model.DataListActionResult;
import org.joget.apps.form.model.FormRow;
import org.joget.apps.form.model.FormRowSet;
import org.joget.apps.form.service.FileUtil;
import org.joget.commons.util.LogUtil;
import org.joget.workflow.util.WorkflowUtil;

/**
 *
 * @author Maxson
 */
public class DocumentGenerationDatalistAction extends DataListActionDefault {

    private final static String MESSAGE_PATH = "message/form/DocumentGenerationDatalistAction";

    @Override
    public String getName() {
        return "Document Generation Datalist Action";
    }

    @Override
    public String getVersion() {
        return "8.0.0";
    }

    @Override
    public String getDescription() {
        return AppPluginUtil.getMessage("org.joget.marketplace.DocumentGenerationDatalistAction.pluginDesc", getClassName(), MESSAGE_PATH);
    }

    @Override
    public String getLinkLabel() {
        return getPropertyString("label");
    }

    @Override
    public String getHref() {
        return getPropertyString("href");
    }

    @Override
    public String getTarget() {
        return "post";
    }

    @Override
    public String getHrefParam() {
        return getPropertyString("hrefParam");
    }

    @Override
    public String getHrefColumn() {
        String recordIdColumn = getPropertyString("recordIdColumn"); //get column id from configured properties options
        if ("id".equalsIgnoreCase(recordIdColumn) || recordIdColumn.isEmpty()) {
            return getPropertyString("hrefColumn"); //Let system to set the primary key column of the binder
        } else {
            return recordIdColumn;
        }
    }

    @Override
    public String getConfirmation() {
        return getPropertyString("confirmation");
    }

    @Override
    public String getLabel() {
        return AppPluginUtil.getMessage("org.joget.marketplace.DocumentGenerationDatalistAction.pluginLabel", getClassName(), MESSAGE_PATH);
    }

    @Override
    public String getClassName() {
        return getClass().getName();
    }

    @Override
    public String getPropertyOptions() {
        return AppUtil.readPluginResource(getClassName(), "/properties/form/documentGenerationDatalistAction.json", null, true, MESSAGE_PATH);
    }

    //Main Function
    @Override
    public DataListActionResult executeAction(DataList dataList, String[] rowKeys) {

        //Only allow POST
        HttpServletRequest request = WorkflowUtil.getHttpServletRequest();
        if (request != null && !"POST".equalsIgnoreCase(request.getMethod())) {
            return null;
        }

        //Check for submitted rows
        if (rowKeys != null && rowKeys.length > 0) {
            try {
                //Get the HTTP response 
                HttpServletResponse response = WorkflowUtil.getHttpServletResponse();

                if (rowKeys.length == 1) {
                    //Generate single pdf for download
                    generateSingleFile(request, response, rowKeys[0]);
                } else {
                    //Generate zip for all selected files
                    generateMultipleFile(request, response, rowKeys);

                }
            } catch (Exception e) {
                LogUtil.error(getClassName(), e, "Failed to generate word file");
            }
        }
        return null;
    }

    //Apache POI
    //Method to replace placeholder with data in every paragraph
    protected void replacePlaceholderInParagraphs(Map<String, String> dataParams, XWPFDocument xwpfDocument) {
        for (Map.Entry<String, String> entry : dataParams.entrySet()) {
            for (XWPFParagraph paragraph : xwpfDocument.getParagraphs()) {
                         String text = paragraph.getText();
                         if (text != null && !text.isEmpty() && text.contains(entry.getKey())) {
                                text = text.replace("${" + entry.getKey() + "}", entry.getValue());
                                //Find a way to insert back to XWPF
                                //Loop through all XWPFParagraph instances and remove each run 
                                //Ensure pararaph is empty 
                                for (int i = paragraph.getRuns().size() - 1; i >= 0; i--) {
                                    paragraph.removeRun(i);
                                }

                                // Create a new run and set the modified text
                                XWPFRun newRun = paragraph.createRun();
                                newRun.setText(text);
                            }
            }
        }
    }

    //Apache POI
    //Method to replace placeholder with data in every tables
    protected void replacePlaceholderInTables(Map<String, String> dataParams, XWPFDocument xwpfDocument) {
        for (Map.Entry<String, String> entry : dataParams.entrySet()) {
            for (XWPFTable xwpfTable : xwpfDocument.getTables()) {
                for (XWPFTableRow xwpfTableRow : xwpfTable.getRows()) {
                    for (XWPFTableCell xwpfTableCell : xwpfTableRow.getTableCells()) {
                        for (XWPFParagraph xwpfParagraph : xwpfTableCell.getParagraphs()) {
                            String text = xwpfParagraph.getText();
                            if (text != null && !text.isEmpty() && text.contains(entry.getKey())) {
                                text = text.replace("${" + entry.getKey() + "}", entry.getValue());
                            //Find a way to insert back to XWPF
                            // Clear existing runs
                            for (int i = xwpfParagraph.getRuns().size() - 1; i >= 0; i--) {
                                xwpfParagraph.removeRun(i);
                            }
                            
                            // Create a new run and set the modified text
                            XWPFRun newRun = xwpfParagraph.createRun();
                            newRun.setText(text);
                            }
                        }
                    }
                }
            }
        }
    }
    
    
    protected void replaceImageInParagraph(Map<String, String> dataParams, XWPFDocument xwpfDocument, String row) {

    for (Map.Entry<String, String> entry : dataParams.entrySet()) {
        for (XWPFParagraph paragraph : xwpfDocument.getParagraphs()) {
            //Credit to Imran Siddique in solving the "capital letter matching issue" in this method
            String text = paragraph.getText();
            if (text != null && !text.isEmpty() && text.contains(entry.getValue())) {
                if (isImageValue(text)) {   
                    try {                    
                        AppDefinition appDef = AppUtil.getCurrentAppDefinition();
                        String formDef = getPropertyString("formDefId");
                        AppService appService = (AppService) AppUtil.getApplicationContext().getBean("appService");
                        String tableName = appService.getFormTableName(appDef, formDef);
                        File file = FileUtil.getFile(text, tableName, row);
                        
                        //New method
                        FileInputStream fileInputStream = new FileInputStream(file);
                        
                        // Remove existing runs from the paragraph
                        for (int i = paragraph.getRuns().size() - 1; i >= 0; i--) {
                            paragraph.removeRun(i);
                        }
                        

                        // Create a new run and insert the image as a picture
                        XWPFRun newRun = paragraph.createRun();
                        newRun.addPicture(fileInputStream, Document.PICTURE_TYPE_PNG, row + "_image", Units.toEMU(400), Units.toEMU(200));
                        
                        System.out.println(entry.getValue() + " successfully added into new word file");
                        fileInputStream.close();
                    } catch (IOException | InvalidFormatException e) {
                        LogUtil.error(getClassName(), e, "Failed to generate word file");
                    }
                }
            }
        }
    }
}

    
    
    //Replace template image placeholder with row image
    //For Paragraph
    protected void replaceImageInTable(Map<String, String> dataParams, XWPFDocument xwpfDocument, String row){
        
        for (Map.Entry<String, String> entry : dataParams.entrySet()) {
            for (XWPFTable xwpfTable : xwpfDocument.getTables()) {
                for (XWPFTableRow xwpfTableRow : xwpfTable.getRows()) {
                    for (XWPFTableCell xwpfTableCell : xwpfTableRow.getTableCells()) {
                        for (XWPFParagraph xwpfParagraph : xwpfTableCell.getParagraphs()) {
                            String text = xwpfParagraph.getText();
                            if (text != null && !text.isEmpty() && text.contains(entry.getValue())) {
                                    if (isImageValue(entry.getValue())) {
                                    try {
                                        AppDefinition appDef = AppUtil.getCurrentAppDefinition();
                                        String formDef = getPropertyString("formDefId");
                                        AppService appService = (AppService) AppUtil.getApplicationContext().getBean("appService");
                                        String tableName = appService.getFormTableName(appDef, formDef);
                                        File file = FileUtil.getFile(entry.getValue(), tableName, row);

                                        FileInputStream fileInputStream = new FileInputStream(file);


                                        // Remove existing runs from the paragraph
                                        for (int i = xwpfParagraph.getRuns().size() - 1; i >= 0; i--) {
                                            xwpfParagraph.removeRun(i);
                                        }
                                        
                                        //Obtain user input width and height from plugin properties
                                        int width = Integer.parseInt(getPropertyString("imageWidth"));
                                        int height = Integer.parseInt(getPropertyString("imageHeight"));
                                        
                                        // Create a new run and insert the image as a picture
                                        XWPFRun newRun = xwpfParagraph.createRun();
                                        newRun.addPicture(fileInputStream, Document.PICTURE_TYPE_JPEG, row + "_image", Units.toEMU(width), Units.toEMU(height));
                                        System.out.println(entry.getValue() + " successfully added into new word file");
                                        fileInputStream.close();

                                    } catch (IOException | InvalidFormatException e) {
                                        LogUtil.error(getClassName(), e, "Failed to generate word file");
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }
        
    }
    
    //Check if passed value is an image
    private boolean isImageValue(String value) {
        if (value.toLowerCase().endsWith(".jpg") || value.toLowerCase().endsWith(".png") || value.toLowerCase().endsWith(".jpeg")) {
            System.out.println("Image File Found = " + value);
            return true;
        } else {
            return false;
        }
    }

    //Get template file method 
    //Get uploaded template file from plugin properties
    protected File getTempFile() throws IOException {
        //Obtain uploaded file from plugin properties
        //Credit to Imran Siddique in solving "obtain uploaded file from plugin properties issue"
        String fileHashVar = getPropertyString("templateFile");
        String templateFilePath = AppUtil.processHashVariable(fileHashVar, null, null, null);
        Path filePath = Paths.get(templateFilePath);
        String fileName = filePath.getFileName().toString();
        AppDefinition appDef = AppUtil.getCurrentAppDefinition();
        File file = AppResourceUtil.getFile(appDef.getAppId(), String.valueOf(appDef.getVersion()), fileName);
        
            //Validation
            if(file.exists()){
                System.out.println("Uploaded Template File Obtained");
                return file;
            }
            else{
                System.out.println("File not found!");
                return null;
            }
    }

    //Generate word file for single datalist row
    protected void generateSingleFile(HttpServletRequest request, HttpServletResponse response, String row) {

        AppDefinition appDef = AppUtil.getCurrentAppDefinition();
        String formDef = getPropertyString("formDefId");
        AppService appService = (AppService) AppUtil.getApplicationContext().getBean("appService");
        //To get whole row of the datalist
        
        //Credit to Hugo in giving guidance on obtaining form row
        FormRowSet formRowSet = appService.loadFormData(appDef.getAppId(), appDef.getVersion().toString(), formDef, row);

        try {
            //Get uploaded template file from plugin properties 
            File tempFile = getTempFile();
            InputStream fInputStream = new FileInputStream(tempFile);

            //Create a XWPFDocument object
            XWPFDocument apachDoc = new XWPFDocument(fInputStream);
            XWPFWordExtractor extractor = new XWPFWordExtractor(apachDoc);
            
            //Extracted Text stored in a String 
            String text = extractor.getText();
            extractor.close();

            //File Text Array & ArrayList (After regex)
            String[] textArr;
            ArrayList<String> textArrayList = new ArrayList<String>();

            //Remove all whitespaces in extracted text
            textArr = text.split("\\s+");
            for (String x : textArr) {
                if (x.startsWith("${") && x.endsWith("}")) {
                    textArrayList.add(x.substring(2, x.length() - 1));
                }
            }

            //Perform Matching Operation
            Map<String, String> matchedMap = new HashMap<String, String>();
            if (formRowSet != null && !formRowSet.isEmpty()) {
                for (String key : textArrayList) {
                    for (FormRow r : formRowSet) {
                        //The keyset of the formrow
                        Set<Object> formSet = r.keySet();

                        //Matching operation => Check if form key match with template key
                        for (Object formKey : formSet) {
                            if (formKey.toString().equals(key)) {
                                matchedMap.put(formKey.toString(), r.getProperty(key));
                            }
                        }

                    }
                }
            }
            
            //Methods to replace all selected datalist data field with template file placeholder variables respectively
            replacePlaceholderInParagraphs(matchedMap, apachDoc);
            replacePlaceholderInTables(matchedMap, apachDoc);
            replaceImageInParagraph(matchedMap, apachDoc, row);
            replaceImageInTable(matchedMap, apachDoc, row);
            
            writeResponseSingle(request, response, apachDoc, row + ".docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
        
        } catch (Exception e) {
            LogUtil.error(this.getClassName(), e, e.toString());
        }
    }

    //Generate word file for multiple datalist row
    protected void generateMultipleFile(HttpServletRequest request, HttpServletResponse response, String[] rows) throws IOException, ServletException {

        //ArrayList of XWPFDocument
        ArrayList<XWPFDocument> documents = new ArrayList<XWPFDocument>();

        for (String row : rows) {
            AppDefinition appDef = AppUtil.getCurrentAppDefinition();
            String formDef = getPropertyString("formDefId");
            AppService appService = (AppService) AppUtil.getApplicationContext().getBean("appService");
            //To get whole row of the datalist
            FormRowSet formRowSet = appService.loadFormData(appDef.getAppId(), appDef.getVersion().toString(), formDef, row);

            try {

                File tempFile = getTempFile();
                InputStream fInputStream = new FileInputStream(tempFile);

                //XWPFDocument 
                XWPFDocument apachDoc = new XWPFDocument(fInputStream);
                XWPFWordExtractor extractor = new XWPFWordExtractor(apachDoc);

                //Extracted Text stored in String 
                String text = extractor.getText();
                extractor.close();

                //File Text Array & ArrayList (After regex)
                String[] textArr;
                ArrayList<String> textArrayList = new ArrayList<String>();

                //Remove all whitespaces in extracted text
                textArr = text.split("\\s+");
                for (String x : textArr) {
                    if (x.startsWith("${") && x.endsWith("}")) {
                        textArrayList.add(x.substring(2, x.length() - 1));
                    }
                }

                //Matching Operation
                Map<String, String> matchedMap = new HashMap<String, String>();
                if (formRowSet != null && !formRowSet.isEmpty()) {
                    for (String key : textArrayList) {
                        //Null => What's the issue 
                        for (FormRow r : formRowSet) {
                            //The keyset of the formrow
                            Set<Object> formSet = r.keySet();

                            //Matching operation => Check if form key match with template key
                            for (Object formKey : formSet) {
                                if (formKey.toString().equals(key)) {
                                    matchedMap.put(formKey.toString(), r.getProperty(key));
                                }
                            }
                        }
                    }
                }
                //Methods to replace all selected datalist data field with template file placeholder variables respectively
                replacePlaceholderInParagraphs(matchedMap, apachDoc);
                replacePlaceholderInTables(matchedMap, apachDoc);
                replaceImageInParagraph(matchedMap, apachDoc, row);
                replaceImageInTable(matchedMap, apachDoc, row);
                //Add the XWPFDocument to the arraylist
                documents.add(apachDoc);
            } catch (Exception e) {
                LogUtil.error(this.getClassName(), e, e.toString());
            }
        }

        writeResponseMulti(request, response, documents , rows);

    }
    
    //Reponse for multi-selection
    protected void writeResponseMulti(HttpServletRequest request, HttpServletResponse response, ArrayList<XWPFDocument> apachDocs , String[] rows) throws IOException, ServletException {

        //Set reponse content type to ZIP
        response.setContentType("application/zip");

        //Set reponse header to specify ZIP file name
        response.setHeader("Content-Disposition", "attachment; filename=WordFiles.zip");

        try ( ZipOutputStream zipOutputStream = new ZipOutputStream(response.getOutputStream())) {
            for (int i = 0; i < apachDocs.size(); i++) {
                //Create a new entry in ZIP file
                ZipEntry zipEntry = new ZipEntry(rows[i] + ".docx");
                zipOutputStream.putNextEntry(zipEntry);

                //Write XWPFDocument to ZIP file
                apachDocs.get(i).write(zipOutputStream);

                zipOutputStream.closeEntry();
            }
            zipOutputStream.flush();

        } catch (Exception e) {
            LogUtil.error(this.getClassName(), e, e.toString());
        }

    }
    
    //Reponse for single selection 
    protected void writeResponseSingle(HttpServletRequest request, HttpServletResponse response, XWPFDocument apachDoc, String fileName, String contentType) throws IOException, ServletException {
        //Get servlet output stream from the response
        ServletOutputStream outputStream = response.getOutputStream();

        try {

            String name = URLEncoder.encode(fileName, "UTF8").replaceAll("\\+", "%20");
            //Set response header to specify the file name 
            response.setHeader("Content-Disposition", "attachment;filename=" + name + ";filename*=UTF-8''" + name);
            //Set response content type
            response.setContentType(contentType + "; charset=UTF-8");

            apachDoc.write(outputStream);

        } finally {

            apachDoc.close();
            outputStream.flush();

        }
    }

}

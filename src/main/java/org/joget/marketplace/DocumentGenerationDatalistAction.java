/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Class.java to edit this template
 */
package org.joget.marketplace;
    
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.net.URI;
import java.net.URL;
import java.net.URLEncoder;
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
import org.apache.poi.xwpf.extractor.XWPFWordExtractor;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.joget.apps.app.model.AppDefinition;
import org.joget.apps.app.service.AppPluginUtil;
import org.joget.apps.app.service.AppService;
import org.joget.apps.app.service.AppUtil;
import org.joget.apps.datalist.model.DataList;
import org.joget.apps.datalist.model.DataListActionDefault;
import org.joget.apps.datalist.model.DataListActionResult;
import org.joget.apps.form.model.FormRow;
import org.joget.apps.form.model.FormRowSet;
import org.joget.apps.form.service.FileUtil;
import org.joget.commons.util.FileManager;
import org.joget.commons.util.LogUtil;
import org.joget.workflow.util.WorkflowUtil;
/**
 *
 * @author Maxson
 */
public class DocumentGenerationDatalistAction extends DataListActionDefault{
    
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
        if(request!=null && !"POST".equalsIgnoreCase(request.getMethod())){
            return null;
        }
        
        //Check for submitted rows
        if(rowKeys !=null && rowKeys.length > 0){
            try{
                //Get the HTTP response 
                HttpServletResponse response = WorkflowUtil.getHttpServletResponse();
                
                if(rowKeys.length == 1){
                    //Generate single pdf for download
                    generateSingleFile(request, response , rowKeys[0]);
                }
                else{
                    //Generate zip for all selected files
                    generateMultipleFile(request, response , rowKeys);
                    
                }
            }catch(Exception e){
                LogUtil.error(getClassName(), e , "Failed to generate word file" );
            }
        }
        return null;
    }

    //NEED TO REVISE (NOPE, STILL STUCK => Proceed to create a form to store template file)
    //Need to find a way to get the template file from plugin properties 
    //Switch to this when we figured it out
    protected File getTemplateFile(){
        
        
        //Trial #1 28/MAY/2023
//        String tempFileName = getPropertyString("templateFile");
//        Object fileObject = getProperty(tempFileName);
//        
//        if(fileObject instanceof Element) {
//            Element fileElement = (Element) fileObject;
//            String uploadedFile = fileElement.getPropertyString("filePath");
//
//            // Create a File object from the uploaded file path
//            File file = new File(uploadedFile);
//            return file;
//        }
        
        //TRIAL #2 28/MAY/2023
//        String tempFileName = getPropertyString("templateFile");
//        AppDefinition appDef = AppUtil.getCurrentAppDefinition();
//        HttpServletRequest request = WorkflowUtil.getHttpServletRequest();
//        String encodedTempFile = tempFileName;
//        
//        try{
//            
//            File tempFile = FileManager.getFileByPath(tempFileName);
//            return tempFile;
//            
//        }catch(Exception e){
//            LogUtil.error(getClassName(), e, "Failed to get uploaded file :(");
//        }
        
        
        //Original Attempt 
        String tempFile = getPropertyString("templateFile");
        AppDefinition appDef = AppUtil.getCurrentAppDefinition();
        HttpServletRequest request = WorkflowUtil.getHttpServletRequest();
        String encodedTempFile = tempFile;
        try{
            //REVISE REVISE
            encodedTempFile =  URLEncoder.encode(encodedTempFile, "UTF8").replaceAll("\\+", "%20");
            String fileURL = "https://" +  request.getServerName() + ":" + request.getServerPort() + request.getContextPath() + "/web/client/app" + appDef.getAppId() + "/" + appDef.getVersion().toString() + "/" + appDef.getName() + "" + "/" ;
//            file = FileUtil.getFile();
        }catch(Exception e){
            LogUtil.error(getClassName(), e, "Failed to get uploaded file!");
        }
        
        return null;
    }
    
    //Apache POI
    //Method to replace placeholder with data in every paragraph
    protected void replacePlaceholderInParagraphs(Map<String, String> dataParams, XWPFDocument xwpfDocument)
    {
        for (Map.Entry<String, String> entry : dataParams.entrySet())
        {
            for (XWPFParagraph paragraph : xwpfDocument.getParagraphs())
            {
                for (XWPFRun run : paragraph.getRuns())
                {
                    String text = run.text();
                    if (
                        text != null
                        && text.contains(entry.getKey())
                        && entry.getValue() != null
                        && !entry.getValue().isEmpty()
                        )
                    {
                       text = text.replace("${" + entry.getKey() + "}", entry.getValue());
                       run.setText(text, 0);
                    }
                }
            }
        }
    }
    
    //Apache POI
    //Method to replace placeholder with data in every tables
    protected void replacePlaceholderInTables(Map<String, String> dataParams, XWPFDocument xwpfDocument)
    {
        for (Map.Entry<String, String> entry : dataParams.entrySet())
        {
            for (XWPFTable xwpfTable : xwpfDocument.getTables())
            {
                for (XWPFTableRow xwpfTableRow : xwpfTable.getRows())
                {
                    for (XWPFTableCell xwpfTableCell : xwpfTableRow.getTableCells())
                    {
                        for (XWPFParagraph xwpfParagraph : xwpfTableCell.getParagraphs())
                        {
                            for (XWPFRun xwpfRun : xwpfParagraph.getRuns())
                            {
                                String text = xwpfRun.text();
                                if (
                                    text != null
                                        && text.contains(entry.getKey())
                                        && entry.getValue() != null
                                        && !entry.getValue().isEmpty()
                                    )
                                {
                                    text = text.replace("${" + entry.getKey() + "}", entry.getValue());
                                    xwpfRun.setText(text, 0);
                                }
                            }
                        }
                    }
                }
            }
        }
    }
    
    //#2 get temp file method (Hardcoded)
    //Temporary method to obtain uploaded file (I'm yet to figure out to how to obtain from 
    protected File getTempFile() throws IOException{
            
        //Hardcoded 
        //Find a way to obtain from file upload plugin properties
        String pKey = "a6af7a75-c330-48be-976c-f14f17928c02";
        String tableName = "toy_tempfile";
        String fileName = getPropertyString("templateFile");

        //The should-be way
        //Obtain uploaded file from plugin properties
        String templateFile = getPropertyString("templateFile");
        String fileDir = FileManager.getBaseDirectory();
        
        File tempFile =  FileUtil.getFile(fileName, tableName, pKey);
        return tempFile;
        

    }
    
    //method that generate single file (single file download)
    protected void  generateSingleFile(HttpServletRequest request, HttpServletResponse response, String row){
        
        AppDefinition appDef = AppUtil.getCurrentAppDefinition(); 
        String formDef = getPropertyString("formDefId");
        AppService appService = (AppService)AppUtil.getApplicationContext().getBean("appService");
        //To get whole row of the datalist
        FormRowSet formRowSet = appService.loadFormData( appDef.getAppId(), appDef.getVersion().toString(), formDef, row);
//      FormRowSet formRowSet = appService.loadFormData(formDef, row);
        
        //Hardcoded placeholder variable (Not used now)
        Map<String , String> pHolder = new HashMap<String, String>();
        pHolder.put("userID", "32cs2@fdi&*@");
        pHolder.put("name", "Champion");
        pHolder.put("desc" , "The Greatest Champion ALIVE");
        pHolder.put("tier" , "SSS");
        
        //Temporary File
        String url = "https://" + request.getServerName() + ":" + request.getServerPort() + "/jw/web/app/toyarchives/resources/TemplateFile.docx";
//      String url =  request.getServerName() + ":" + request.getServerPort() + "/jw/web/app/toyarchives/resources/TemplateFile.docx";
        URI uri ;
        URL templateURL;
        File templateFile;
        
        //(TEMPORARY) Get template file from 
        
        //Placeholder 
        try{
         //Temporary Template File URL
//         templateURL = new URL (url);
         //URL to URI
//         uri = templateURL.toURI();
         //URI to file
//         templateFile = new File(new URL(url).toURI());

         //new URL to file (17-May)
//         BufferedInputStream in = new BufferedInputStream(new URL(url).openStream());


         //start rework here
         //File input stream
//         FileInputStream fInputStream = new FileInputStream(templateFile);
         
        //Get uploaded file from plugin properties (28/MAY/2023)
        //Get file from 
        File tempFile = getTempFile(); 
//        String testFilePath = getPropertyString("templateFile");
        InputStream fInputStream = new FileInputStream(tempFile);
        
        
         //new input stream (Require Rework);
//         InputStream fInputStream = new URL(url).openStream();

         //XWPFDocument 
         XWPFDocument apachDoc = new XWPFDocument(fInputStream);
         XWPFWordExtractor extractor =  new XWPFWordExtractor(apachDoc);
         
         //Extracted Text stored in String 
         String text = extractor.getText();
         extractor.close();
        
         //File Text Array & ArrayList (After regex)
         String [] textArr;
         ArrayList<String> textArrayList = new ArrayList<String>();
         
         //Remove all whitespaces in extracted text
         textArr = text.split("\\s+");
         for(String x:textArr){
             if(x.startsWith("${") && x.endsWith("}")){
                 textArrayList.add(x.substring(2,x.length()-1));
             }
         }
         
         
         
        //Matching Operation
        Map<String , String> matchedMap = new HashMap<String, String>();
                if(formRowSet != null && !formRowSet.isEmpty()){
                    for( String key : textArrayList){
                        //Null => What's the issue 
                        for(FormRow r : formRowSet){
                            //The keyset of the formrow
                            Set<Object> formSet = r.keySet();
                            
                            //Matching operation => Check if form key match with template key
                            for(Object formKey: formSet){
                                if(formKey.toString().equals(key)){
                                    matchedMap.put(formKey.toString() , r.getProperty(key));
                                }
                            }
                   
                        }
                    }    
                }
        //Methods to replace placeholder variables
        //For paragraph, tables => Might need to add more => awaiting review
         replacePlaceholderInParagraphs(matchedMap , apachDoc);
         replacePlaceholderInTables(matchedMap, apachDoc);       
         
        //Find a way to modify this => to allow reusability for this method (generateSingleFile)
         writeResponseSingle(request, response, apachDoc ,row + ".docx" , "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
        }catch(Exception e){
            LogUtil.error(this.getClassName(), e, e.toString());
        }
    }
    
    //Method that generates multiple file for download
    //Put it into zip
    protected void generateMultipleFile(HttpServletRequest request, HttpServletResponse response , String[] rows) throws IOException, ServletException{
        
        //ArrayList of XWPFDocument
        ArrayList<XWPFDocument> documents = new ArrayList<XWPFDocument>(); 
        
        for(String row : rows){
            AppDefinition appDef = AppUtil.getCurrentAppDefinition(); 
            String formDef = getPropertyString("formDefId");
            AppService appService = (AppService)AppUtil.getApplicationContext().getBean("appService");
            //To get whole row of the datalist
            FormRowSet formRowSet = appService.loadFormData( appDef.getAppId(), appDef.getVersion().toString(), formDef, row);

            //Placeholder 
            try{           
            //start rework here
            //Get uploaded file from plugin properties (28/MAY/2023)
            //Get file from 
            File tempFile = getTempFile(); 
            //String testFilePath = getPropertyString("templateFile");
            InputStream fInputStream = new FileInputStream(tempFile);

             //XWPFDocument 
             XWPFDocument apachDoc = new XWPFDocument(fInputStream);
             XWPFWordExtractor extractor =  new XWPFWordExtractor(apachDoc);

             //Extracted Text stored in String 
             String text = extractor.getText();
             extractor.close();

             //File Text Array & ArrayList (After regex)
             String [] textArr;
             ArrayList<String> textArrayList = new ArrayList<String>();

             //Remove all whitespaces in extracted text
             textArr = text.split("\\s+");
             for(String x:textArr){
                 if(x.startsWith("${") && x.endsWith("}")){
                     textArrayList.add(x.substring(2,x.length()-1));
                 }
             }

            //Matching Operation
            Map<String , String> matchedMap = new HashMap<String, String>();
                    if(formRowSet != null && !formRowSet.isEmpty()){
                        for( String key : textArrayList){
                            //Null => What's the issue 
                            for(FormRow r : formRowSet){
                                //The keyset of the formrow
                                Set<Object> formSet = r.keySet();

                                //Matching operation => Check if form key match with template key
                                for(Object formKey: formSet){
                                    if(formKey.toString().equals(key)){
                                        matchedMap.put(formKey.toString() , r.getProperty(key));
                                    }
                                }
                            }
                        }    
                    }
            //Methods to replace placeholder variables
            //For paragraph, tables => Might add more 
             replacePlaceholderInParagraphs(matchedMap , apachDoc);
             replacePlaceholderInTables(matchedMap, apachDoc);
             
             //Add the XWPFDocument to the arraylist
             documents.add(apachDoc);
            }catch(Exception e){
                LogUtil.error(this.getClassName(), e, e.toString());
            }
        }
        
        writeResponseMulti(request , response , documents);
         
    }
    
    protected void writeResponseMulti(HttpServletRequest request , HttpServletResponse response , ArrayList<XWPFDocument> apachDocs)throws IOException, ServletException{
        
        //Set reponse content type to ZIP
        response.setContentType("application/zip");
        
        //Set reponse header to specify ZIP file name
        response.setHeader("Content-Disposition", "attachment; filename=documents.zip");
        
        try(ZipOutputStream zipOutputStream = new ZipOutputStream(response.getOutputStream())){
            for(int i = 0; i<apachDocs.size() ; i++){
                //Create a new entry in ZIP file
                ZipEntry zipEntry = new ZipEntry("File_" + (i+1) + ".docx");
                zipOutputStream.putNextEntry(zipEntry);
                
                //Write XWPFDocument to ZIP file
                apachDocs.get(i).write(zipOutputStream);
                
                zipOutputStream.closeEntry();
                
            }
            zipOutputStream.flush();
            
        }catch(Exception e){
            LogUtil.error(this.getClassName(), e, e.toString());
        }
        
        
        
    }
    
    protected void writeResponseSingle(HttpServletRequest request , HttpServletResponse response, XWPFDocument apachDoc ,  String fileName , String contentType) throws IOException, ServletException{
            //Get servlet output stream from the response
            ServletOutputStream outputStream = response.getOutputStream();
         
        try{
            
            String name = URLEncoder.encode(fileName, "UTF8").replaceAll("\\+", "%20");
            //Set response header to specify the file name 
            response.setHeader("Content-Disposition", "attachment;filename="+name+";filename*=UTF-8''" + name);
            //Set response content type
            response.setContentType(contentType+"; charset=UTF-8");

            apachDoc.write(outputStream);
            

        }finally{
        
               apachDoc.close();
               outputStream.flush();
//               request.getRequestDispatcher(fileName).forward(request, response);
               
        }
    }
    
}

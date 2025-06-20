package org.joget.marketplace.util;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.math.BigInteger;
import java.net.URLEncoder;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.Map;
import java.util.Set;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
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
import org.apache.poi.xwpf.usermodel.XWPFTable.XWPFBorderType;
import org.apache.xmlbeans.XmlCursor;
import org.joget.apps.app.dao.FormDefinitionDao;
import org.joget.apps.app.model.AppDefinition;
import org.joget.apps.app.model.FormDefinition;
import org.joget.apps.app.service.AppResourceUtil;
import org.joget.apps.app.service.AppService;
import org.joget.apps.app.service.AppUtil;
import org.joget.apps.form.model.FormRow;
import org.joget.apps.form.model.FormRowSet;
import org.joget.apps.form.service.FileUtil;
import org.joget.commons.util.LogUtil;
import org.json.JSONArray;
import org.json.JSONObject;

import com.google.gson.JsonArray;
import com.google.gson.JsonObject;
import com.google.gson.JsonParser;
import java.util.List;
import org.joget.apps.form.service.FormUtil;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblGrid;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblWidth;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTcPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STJc;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STTblWidth;

public class DocumentGenerationUtil {

    private static File generatedFile;

      protected static void replacePlaceholderInParagraphs(Map<String, String> dataParams, XWPFDocument xwpfDocument, String formDefId, String gridIncludeHeader, String gridDirection, String gridWidth) {
        for (Map.Entry<String, String> entry : dataParams.entrySet()) {
            for (XWPFParagraph paragraph : new ArrayList<>(xwpfDocument.getParagraphs())) {
                String text = paragraph.getText();
                if (text != null && !text.isEmpty() && text.contains(entry.getKey())) {
                    text = text.replace("${" + entry.getKey() + "}", entry.getValue());
                    for (int i = paragraph.getRuns().size() - 1; i >= 0; i--) {
                        paragraph.removeRun(i);
                    }

                    // if value is json
                    if (text.contains("[") || text.contains("]")) {
                        int start = text.indexOf('[');
                        int end = text.lastIndexOf(']') + 1;

                        String label = text.substring(0, start).trim();
                        String jsonPart = text.substring(start, end);
                        String endLabel = "";
                        if (end < text.length()) {
                            endLabel = text.substring(end).trim();
                        }

                        for (int i = paragraph.getRuns().size() - 1; i >= 0; i--) {
                            paragraph.removeRun(i);
                        }
                        XWPFParagraph labelParagraph = xwpfDocument.insertNewParagraph(paragraph.getCTP().newCursor());
                        XWPFRun labelRun = labelParagraph.createRun();
                        labelRun.setText(label);

                        replacePlaceholderInJSON(entry.getKey(), jsonPart, xwpfDocument, labelParagraph, formDefId, gridIncludeHeader, gridDirection, gridWidth);
                    
                        XWPFParagraph endLabelParagraph = xwpfDocument.insertNewParagraph(paragraph.getCTP().newCursor());
                        XWPFRun endLabelRun = endLabelParagraph.createRun();
                        endLabelRun.setText(endLabel);
                    } else {
                        XWPFRun newRun = paragraph.createRun();
                        newRun.setText(text);
                    }
                }
            }
        }
    }

    protected static void replacePlaceholderInJSON(String textKey, String text, XWPFDocument xwpfDocument, XWPFParagraph paragraph, String formDefId, String gridIncludeHeader, String gridDirection, String gridWidth) {
        AppDefinition appDef = AppUtil.getCurrentAppDefinition();
        String formDef = formDefId;
        FormDefinitionDao formDefinitionDao = (FormDefinitionDao) FormUtil.getApplicationContext().getBean("formDefinitionDao");
        FormDefinition formDefinition = formDefinitionDao.loadById(formDef, appDef);
    
        LinkedHashMap<String, String> headerMap = new LinkedHashMap<>();
        if (formDefinition != null) {
            JSONObject rootObject = new JSONObject(formDefinition.getJson());
            extractHeaderMapFromGrid(textKey, rootObject, headerMap);
        }
        JsonArray jsonArray = JsonParser.parseString(text).getAsJsonArray();
    
        boolean includeHeader = "true".equals(gridIncludeHeader);
        List<String> orderedKeys = new ArrayList<>(headerMap.keySet());
        List<List<String>> tableData = new ArrayList<>();
    
        for (int i = 0; i < jsonArray.size(); i++) {
            JsonObject jsonObject = jsonArray.get(i).getAsJsonObject();
            List<String> rowValues = new ArrayList<>();
    
            for (String key : orderedKeys) {
                rowValues.add(jsonObject.has(key) ? jsonObject.get(key).getAsString() : "");
            }
    
            tableData.add(rowValues);
        }
    
        // create table
        int dataRowCount = tableData.size();
        int colCount = orderedKeys.size();

        int actualRows = "vertical".equals(gridDirection)
                ? dataRowCount + (includeHeader ? 1 : 0)
                : colCount;
        int actualCols = "vertical".equals(gridDirection)
                ? colCount
                : dataRowCount + (includeHeader ? 1 : 0);

        XWPFTable table = createEmptyGridTable(actualRows, actualCols, xwpfDocument, paragraph, gridWidth);

        // insert headers
        if (includeHeader) {
            for (int i = 0; i < orderedKeys.size(); i++) {
                String headerLabel = headerMap.getOrDefault(orderedKeys.get(i), orderedKeys.get(i));
                if ("vertical".equals(gridDirection)) {
                    table.getRow(0).getCell(i).setText(headerLabel);
                } else {
                    table.getRow(i).getCell(0).setText(headerLabel);
                }
            }
        }

        // insert data rows
        int dataStartIndex = includeHeader ? 1 : 0;
        for (int rowIdx = 0; rowIdx < tableData.size(); rowIdx++) {
            List<String> row = tableData.get(rowIdx);
            for (int colIdx = 0; colIdx < row.size(); colIdx++) {
                String value = row.get(colIdx);
                if ("vertical".equals(gridDirection)) {
                    table.getRow(dataStartIndex + rowIdx).getCell(colIdx).setText(value);
                } else {
                    table.getRow(colIdx).getCell(dataStartIndex + rowIdx).setText(value);
                }
            }
        }
    }

    protected static void replacePlaceholderInTables(Map<String, String> dataParams, XWPFDocument xwpfDocument) {
        for (Map.Entry<String, String> entry : dataParams.entrySet()) {
            for (XWPFTable xwpfTable : xwpfDocument.getTables()) {
                xwpfTable.getCTTbl().getTblPr().addNewJc().setVal(STJc.CENTER);

                for (XWPFTableRow xwpfTableRow : xwpfTable.getRows()) {
                    for (XWPFTableCell xwpfTableCell : xwpfTableRow.getTableCells()) {
                        for (XWPFParagraph xwpfParagraph : xwpfTableCell.getParagraphs()) {
                            String text = xwpfParagraph.getText();
                            if (text != null && !text.isEmpty() && text.contains(entry.getKey())) {
                                text = text.replace("${" + entry.getKey() + "}", entry.getValue());
                                for (int i = xwpfParagraph.getRuns().size() - 1; i >= 0; i--) {
                                    xwpfParagraph.removeRun(i);
                                }
                                XWPFRun newRun = xwpfParagraph.createRun();
                                newRun.setText(text);
                            }
                        }
                    }
                }
            }
        }
    }

    protected static void extractHeaderMapFromGrid(String textKey, JSONObject jsonObject, LinkedHashMap<String, String> headerMap) {
        if (jsonObject.has("elements") && jsonObject.get("elements") instanceof JSONArray) {
            JSONArray elementsArray = jsonObject.getJSONArray("elements");
    
            for (int i = 0; i < elementsArray.length(); i++) {
                JSONObject element = elementsArray.getJSONObject(i);
    
                if (element.has("className")) {
                    String className = element.getString("className");

                    // only for grid types
                    if (className.equals("org.joget.plugin.enterprise.AdvancedGrid") ||
                        className.equals("org.joget.plugin.enterprise.FormGrid") ||
                        className.equals("org.joget.plugin.enterprise.ListGrid") ||
                         className.equals("org.joget.apps.form.lib.Grid")) {
    
                        if (element.has("properties")) {
                            JSONObject properties = element.getJSONObject("properties");
                            if (properties.has("id") && properties.getString("id").equals(textKey)) {
                                if (properties.has("options")) {
                                    Object options = properties.get("options");
                                    if (options instanceof JSONArray) {
                                        JSONArray optionsArray = (JSONArray) options;
                                        for (int j = 0; j < optionsArray.length(); j++) {
                                            JSONObject option = optionsArray.getJSONObject(j);
                                            if (option.has("value") && option.has("label")) {
                                                headerMap.put(option.getString("value"), option.getString("label"));
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }

                extractHeaderMapFromGrid(textKey, element, headerMap);
            }
        }
    }
    
    protected static void fixPreCreatedTableFormatting(XWPFDocument xwpfDocument, String gridWidthStr) {
        for (XWPFTable table : xwpfDocument.getTables()) {
            // Ensure center alignment
            table.getCTTbl().getTblPr().addNewJc().setVal(STJc.CENTER);

            // Adjust column widths dynamically
            int numCols = table.getRow(0).getTableCells().size();
            if (gridWidthStr != "") {
                int gridWidth = Integer.parseInt(gridWidthStr);
                BigInteger columnWidth = BigInteger.valueOf(gridWidth / numCols);
    
                // Define table grid
                CTTblGrid tblGrid = table.getCTTbl().addNewTblGrid();
                for (int i = 0; i < numCols; i++) {
                    tblGrid.addNewGridCol().setW(columnWidth);
                }
    
                // Apply column width to each cell
                for (XWPFTableRow row : table.getRows()) {
                    for (XWPFTableCell cell : row.getTableCells()) {
                        CTTcPr tcPr = cell.getCTTc().addNewTcPr();
                        CTTblWidth cellWidth = tcPr.addNewTcW();
                        cellWidth.setW(columnWidth);
                        cellWidth.setType(STTblWidth.DXA);
                    }
                }
            }
        }
    }

    protected static void replaceImageInParagraph(Map<String, String> dataParams, XWPFDocument xwpfDocument, String row, String formDefId) {

        for (Map.Entry<String, String> entry : dataParams.entrySet()) {
            for (XWPFParagraph paragraph : xwpfDocument.getParagraphs()) {
                String text = paragraph.getText();
                if (text != null && !text.isEmpty() && text.contains(entry.getValue())) {
                    if (isImageValue(text)) {
                        try {
                            AppDefinition appDef = AppUtil.getCurrentAppDefinition();
                            String formDef = formDefId;
                            AppService appService = (AppService) AppUtil.getApplicationContext().getBean("appService");
                            String tableName = appService.getFormTableName(appDef, formDef);
                            File file = FileUtil.getFile(text, tableName, row);
                            FileInputStream fileInputStream = new FileInputStream(file);
                            for (int i = paragraph.getRuns().size() - 1; i >= 0; i--) {
                                paragraph.removeRun(i);
                            }
                            XWPFRun newRun = paragraph.createRun();
                            newRun.addPicture(fileInputStream, Document.PICTURE_TYPE_PNG, row + "_image", Units.toEMU(400), Units.toEMU(200));
                            fileInputStream.close();
                        } catch (IOException | InvalidFormatException e) {
                            LogUtil.error(getClassName(), e, "Failed to generate word file");
                        }
                    }
                }
            }
        }
    }

    protected static void replaceImageInTable(Map<String, String> dataParams, XWPFDocument xwpfDocument, String row, String formDefId, String imageWidth, String imageHeight) {
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
                                        String formDef = formDefId;
                                        AppService appService = (AppService) AppUtil.getApplicationContext().getBean("appService");
                                        String tableName = appService.getFormTableName(appDef, formDef);
                                        File file = FileUtil.getFile(entry.getValue(), tableName, row);
                                        FileInputStream fileInputStream = new FileInputStream(file);
                                        for (int i = xwpfParagraph.getRuns().size() - 1; i >= 0; i--) {
                                            xwpfParagraph.removeRun(i);
                                        }
                                        int width = Integer.parseInt(imageWidth);
                                        int height = Integer.parseInt(imageHeight);

                                        XWPFRun newRun = xwpfParagraph.createRun();
                                        newRun.addPicture(fileInputStream, Document.PICTURE_TYPE_JPEG, row + "_image", Units.toEMU(width), Units.toEMU(height));
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

    private static boolean isImageValue(String value) {
        if (value.toLowerCase().endsWith(".jpg") || value.toLowerCase().endsWith(".png") || value.toLowerCase().endsWith(".jpeg")) {
            return true;
        } else {
            return false;
        }
    }

    protected static XWPFTable createEmptyGridTable(int rows, int cols, XWPFDocument xwpfDocument, XWPFParagraph paragraph, String gridWidthStr) {
        XmlCursor cursor = paragraph.getCTP().newCursor();
        cursor.toNextSibling();
        XWPFTable table = xwpfDocument.insertNewTbl(cursor);

        table.getCTTbl().getTblPr().addNewJc().setVal(STJc.CENTER);

        // Dynamically set grid width based on user property
        int gridWidth = Integer.parseInt(gridWidthStr);
        BigInteger columnWidth = BigInteger.valueOf(gridWidth);
        CTTblGrid tblGrid = table.getCTTbl().addNewTblGrid();
        for (int i = 0; i < cols; i++) {
            tblGrid.addNewGridCol().setW(columnWidth);
        }

        // Create rows and columns dynamically
        for (int i = 0; i < rows; i++) {
            XWPFTableRow row = (i == 0) ? table.getRow(0) : table.createRow();
            for (int j = 0; j < cols; j++) {
                if (row.getTableCells().size() <= j) {
                    row.createCell();
                }
            }
        }

        // Set margins inside the table cells for better spacing
        for (XWPFTableRow row : table.getRows()) {
            for (XWPFTableCell cell : row.getTableCells()) {
                CTTcPr tcPr = cell.getCTTc().addNewTcPr();
                CTTblWidth cellWidth = tcPr.addNewTcW();
                cellWidth.setW(columnWidth);
                cellWidth.setType(STTblWidth.DXA);
            }
        }

        table.setInsideHBorder(XWPFBorderType.THICK, 5, 0, "000000");
        table.setInsideVBorder(XWPFBorderType.THICK, 5, 0, "000000");
        table.setTopBorder(XWPFBorderType.THICK, 5, 0, "000000");
        table.setBottomBorder(XWPFBorderType.THICK, 5, 0, "000000");
        table.setLeftBorder(XWPFBorderType.THICK, 5, 0, "000000");
        table.setRightBorder(XWPFBorderType.THICK, 5, 0, "000000");

        return table;
    }

    protected static File getTempFile(String templateFile) throws IOException {
        String fileHashVar = templateFile;
        String templateFilePath = AppUtil.processHashVariable(fileHashVar, null, null, null);
        Path filePath = Paths.get(templateFilePath);
        String fileName = filePath.getFileName().toString();
        AppDefinition appDef = AppUtil.getCurrentAppDefinition();
        File file = AppResourceUtil.getFile(appDef.getAppId(), String.valueOf(appDef.getVersion()), fileName);
        //Validation
        if (file.exists()) {
            return file;
        } else {
            return null;
        }
    }

    public static void generateSingleFile(HttpServletRequest request, HttpServletResponse response, String row, String formDefId, String templateFile, String gridIncludeHeader, String gridDirection, String wordFileName, String gridWidth, String imageWidth, String imageHeight, String fileOutput) {

        AppDefinition appDef = AppUtil.getCurrentAppDefinition();
        String formDef = formDefId;

        // Get form data with subform information
        Map<String, Object> formData = FormUtil.loadFormData(
                appDef.getAppId(),
                appDef.getVersion().toString(),
                formDef,
                row,
                true, // includeSubformData
                true, // includeReferenceElements
                false, // flatten
                null // no workflow assignment
        );

        // Convert formData to array format
        FormRowSet formDataRowSet = new FormRowSet();
        FormRow dataRow = new FormRow();
        // Handle each entry in the form data
        for (Map.Entry<String, Object> entry : formData.entrySet()) {
            String key = entry.getKey();
            Object value = entry.getValue();

            if (value instanceof List) {
                // Convert HashMap structure to JSON string
                String jsonValue = convertHashMapToJson(value);
                dataRow.put(key, jsonValue);
            } else {
                dataRow.put(key, value != null ? value.toString() : "");
            }
        }
        formDataRowSet.add(dataRow);

        try {
            File tempFile = getTempFile(templateFile);
            InputStream fInputStream = Files.newInputStream(tempFile.toPath());

            //Create a XWPFDocument object
            XWPFDocument apachDoc = new XWPFDocument(fInputStream);
            fixPreCreatedTableFormatting(apachDoc, gridWidth);

            XWPFWordExtractor extractor = new XWPFWordExtractor(apachDoc);
            String text = extractor.getText();
            extractor.close();

            ArrayList<String> textArrayList = new ArrayList<>();
            String[] textArr = text.split("\\s+");
            for (String x : textArr) {
                if (x.startsWith("${") && x.endsWith("}")) {
                    textArrayList.add(x.substring(2, x.length() - 1));
                }
            }

            //Perform Matching Operation
            Map<String, String> matchedMap = new HashMap<>();
            if (!formDataRowSet.isEmpty()) {
                for (String key : textArrayList) {
                    for (FormRow r : formDataRowSet) {
                        //The keyset of the formrow
                        Set<Object> formSet = r.keySet();

                        //Matching operation => Check if form key match with template key
                        for (Object formKey : formSet) {
                            //if text follows format "json[1].jsonKey", translate json array format
                            Pattern pattern = Pattern.compile("([a-zA-Z]+)\\[(\\d+)]\\.(.+)");
                            Matcher matcher = pattern.matcher(key);

                            if (matcher.matches()) {
                                String jsonName = matcher.group(1);
                                String rowNum = matcher.group(2);
                                String jsonKey = matcher.group(3);
                                if (formKey.toString().equals(jsonName)) {
                                    String jsonString = r.getProperty(jsonName);

                                    JSONArray jsonArray = new JSONArray(jsonString);

                                    if (jsonArray.length() > Integer.parseInt(rowNum)) {
                                        JSONObject jsonObject = jsonArray.getJSONObject(Integer.parseInt(rowNum));
                                        String jsonValue = jsonObject.getString(jsonKey);
                                        matchedMap.put(key, jsonValue);
                                    }
                                }
                            }

                            if (formKey.toString().equals(key)) {
                                String value = r.getProperty(key);
                                matchedMap.put(formKey.toString(), r.getProperty(key));
                            }
                        }
                    }
                }
            }

            //Methods to replace all selected datalist data field with template file placeholder variables respectively
            replacePlaceholderInParagraphs(matchedMap, apachDoc, formDefId, gridIncludeHeader, gridDirection, gridWidth);
            replacePlaceholderInTables(matchedMap, apachDoc);
            replaceImageInParagraph(matchedMap, apachDoc, row, formDefId);
            replaceImageInTable(matchedMap, apachDoc, row, formDefId, imageWidth, imageHeight);

            String customFileName = wordFileName;
            if (customFileName == null || customFileName.isEmpty()) {
                customFileName = "Doc File";
            }
            customFileName = customFileName.replace("{row}", row) + ".docx"; 

            if (fileOutput.equals("file")){
                generatedFile = generateOutputFile(apachDoc, customFileName);
            } else {
                writeResponseSingle(request, response, apachDoc, customFileName);
            }
        } catch (Exception e) {
            LogUtil.error(getClassName(), e, e.toString());
        }
    }

    //Generate word file for multiple datalist row
    public static void generateMultipleFile(HttpServletRequest request, HttpServletResponse response, String[] rows, String formDefId, String templateFile, String gridIncludeHeader, String gridDirection, String zipFileName, String gridWidth, String imageWidth, String imageHeight) throws IOException, ServletException {

        //ArrayList of XWPFDocument
        ArrayList<XWPFDocument> documents = new ArrayList<>();

        for (String row : rows) {
            AppDefinition appDef = AppUtil.getCurrentAppDefinition();
            String formDef = formDefId;

            // Get form data with subform information
            Map<String, Object> formData = FormUtil.loadFormData(
                    appDef.getAppId(),
                    appDef.getVersion().toString(),
                    formDef,
                    row,
                    true, // includeSubformData
                    true, // includeReferenceElements
                    false, // flatten
                    null // no workflow assignment
            );

            // Convert formData to array format
            FormRowSet formDataRowSet = new FormRowSet();
            FormRow dataRow = new FormRow();
            // Handle each entry in the form data
            for (Map.Entry<String, Object> entry : formData.entrySet()) {
                String key = entry.getKey();
                Object value = entry.getValue();

                if (value instanceof List) {
                    // Convert HashMap structure to JSON string
                    String jsonValue = convertHashMapToJson(value);
                    dataRow.put(key, jsonValue);
                } else {
                    dataRow.put(key, value != null ? value.toString() : "");
                }
            }
            formDataRowSet.add(dataRow);

            try {
                File tempFile = getTempFile(templateFile);
                InputStream fInputStream = Files.newInputStream(tempFile.toPath());

                // XWPFDocument 
                XWPFDocument apachDoc = new XWPFDocument(fInputStream);
                fixPreCreatedTableFormatting(apachDoc, gridWidth);

                ArrayList<String> textArrayList = getStrings(apachDoc);

                // Matching Operation
                Map<String, String> matchedMap = new HashMap<>();
                for (String key : textArrayList) {
                    for (FormRow r : formDataRowSet) {
                        Set<Object> formSet = r.keySet();
                        for (Object formKey : formSet) {
                            //if text follows format "json[1].jsonKey", translate json array format
                            Pattern pattern = Pattern.compile("([a-zA-Z]+)\\[(\\d+)]\\.(.+)");
                            Matcher matcher = pattern.matcher(key);

                            if (matcher.matches()) {
                                String jsonName = matcher.group(1);
                                String rowNum = matcher.group(2);
                                String jsonKey = matcher.group(3);

                                if (formKey.toString().equals(jsonName)) {
                                    String jsonString = r.getProperty(jsonName);
                                    JSONArray jsonArray = new JSONArray(jsonString);

                                    if (jsonArray.length() > Integer.parseInt(rowNum)) {
                                        JSONObject jsonObject = jsonArray.getJSONObject(Integer.parseInt(rowNum));
                                        String jsonValue = jsonObject.getString(jsonKey);
                                        matchedMap.put(key, jsonValue);
                                    }
                                }
                            }

                            if (formKey.toString().equals(key)) {
                                // String value = r.getProperty(key);
                                matchedMap.put(formKey.toString(), r.getProperty(key));
                            }
                        }
                    }
                }
                replacePlaceholderInParagraphs(matchedMap, apachDoc, formDefId, gridIncludeHeader, gridDirection, gridWidth);
                replacePlaceholderInTables(matchedMap, apachDoc);
                replaceImageInParagraph(matchedMap, apachDoc, row, formDefId);
                replaceImageInTable(matchedMap, apachDoc, row, formDefId, imageWidth, imageHeight);
                documents.add(apachDoc);
            } catch (Exception e) {
                LogUtil.error(getClassName(), e, e.toString());
            }
        }
        writeResponseMulti(request, response, documents, rows, zipFileName);
    }

    private static ArrayList<String> getStrings(XWPFDocument apachDoc) throws IOException {
        XWPFWordExtractor extractor = new XWPFWordExtractor(apachDoc);

        // Extracted Text stored in String
        String text = extractor.getText();
        extractor.close();

        // File Text Array & ArrayList (After regex)
        ArrayList<String> textArrayList = new ArrayList<>();
        String[] textArr = text.split("\\s+");
        for (String x : textArr) {
            if (x.startsWith("${") && x.endsWith("}")) {
                textArrayList.add(x.substring(2, x.length() - 1));
            }
        }
        return textArrayList;
    }

    protected static void writeResponseMulti(HttpServletRequest request, HttpServletResponse response, ArrayList<XWPFDocument> apachDocs, String[] rows, String zipFileName) throws IOException, ServletException {
        response.setContentType("application/zip");
        String customZipName = zipFileName;
        if (customZipName == null || customZipName.isEmpty()) {
            customZipName = "Wordfile.zip";
        } else {
            customZipName = customZipName.replace("{timestamp}", String.valueOf(System.currentTimeMillis())) + ".zip";
        }

        response.setHeader("Content-Disposition", "attachment; filename=" + customZipName);
        try (ZipOutputStream zipOutputStream = new ZipOutputStream(response.getOutputStream())) {
            for (int i = 0; i < apachDocs.size(); i++) {
                ZipEntry zipEntry = new ZipEntry(rows[i] + ".docx");
                zipOutputStream.putNextEntry(zipEntry);
                apachDocs.get(i).write(zipOutputStream);
                zipOutputStream.closeEntry();
            }
            zipOutputStream.flush();

        } catch (Exception e) {
            LogUtil.error(getClassName(), e, e.toString());
        }
    }

    protected static void writeResponseSingle(HttpServletRequest request, HttpServletResponse response, XWPFDocument apachDoc, String fileName) throws IOException, ServletException {
        ServletOutputStream outputStream = response.getOutputStream();
        try {
            String name = URLEncoder.encode(fileName, "UTF8").replaceAll("\\+", "%20");
            response.setHeader("Content-Disposition", "attachment;filename=" + name + ";filename*=UTF-8''" + name);
            response.setContentType("application/vnd.openxmlformats-officedocument.wordprocessingml.document" + "; charset=UTF-8");
            apachDoc.write(outputStream);

        } finally {
            apachDoc.close();
            outputStream.flush();
            outputStream.close();
        }
    }
    
    protected static String convertHashMapToJson(Object data) {
        if (data instanceof List) {
            try {
                // Convert the List<Map> structure to a JSONArray
                JSONArray jsonArray = new JSONArray();
                List<?> dataList = (List<?>) data;

                for (Object item : dataList) {
                    if (item instanceof Map) {
                        JSONObject jsonObject = new JSONObject((Map<?, ?>) item);
                        jsonArray.put(jsonObject);
                    }
                }

                return jsonArray.toString();
            } catch (Exception e) {
                LogUtil.error(getClassName(), e, "Error converting List to JSON");
            }
        }

        return data.toString();
    }

    public static String getClassName() {
        return "DocumentGenerationUtil";
    }

    protected static File generateOutputFile(XWPFDocument apachDoc, String fileName) throws IOException {
          File outFile = new File(fileName);

        try (FileOutputStream out = new FileOutputStream(outFile)) {
            apachDoc.write(out);
        }

        return outFile;
    }

    public static File getGeneratedFile() {
        return generatedFile;
    }

}

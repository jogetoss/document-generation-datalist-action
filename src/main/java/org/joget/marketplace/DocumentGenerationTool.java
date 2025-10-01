package org.joget.marketplace;

import java.io.File;
import java.io.IOException;
import java.nio.file.Files;
import java.util.Map;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.joget.apps.app.model.AppDefinition;
import org.joget.apps.app.service.AppPluginUtil;
import org.joget.apps.app.service.AppService;
import org.joget.apps.app.service.AppUtil;
import org.joget.apps.form.model.FormRow;
import org.joget.apps.form.model.FormRowSet;
import org.joget.apps.form.service.FileUtil;
import org.joget.apps.form.service.FormUtil;
import org.joget.commons.util.LogUtil;
import org.joget.marketplace.util.DocumentGenerationUtil;
import org.joget.plugin.base.DefaultApplicationPlugin;
import org.joget.workflow.model.WorkflowAssignment;
import org.joget.workflow.util.WorkflowUtil;

public class DocumentGenerationTool extends DefaultApplicationPlugin {

    private final static String MESSAGE_PATH = "message/DocumentGenerationTool";

    @Override
    public Object execute(Map properties) {

        HttpServletRequest request = WorkflowUtil.getHttpServletRequest();
        if (request != null && !"POST".equalsIgnoreCase(request.getMethod())) {
            return null;
        }
        HttpServletResponse response = WorkflowUtil.getHttpServletResponse();

        AppService appService = (AppService) FormUtil.getApplicationContext().getBean("appService");
        AppDefinition appDef = (AppDefinition) properties.get("appDef");
        WorkflowAssignment wfAssignment = (WorkflowAssignment) properties.get("workflowAssignment");
        String recordId = getPropertyString("recordId");
        if (recordId.equals("")) {
            recordId = appService.getOriginProcessId(wfAssignment.getProcessId());
        }

        String filePath = getPropertyString("filePath");;
        String formDefId = getPropertyString("formDefId");
        String fileFieldId = getPropertyString("fileFieldId");
        String pathOptions = getPropertyString("pathOptions");
        String templateFile = getPropertyString("templateFile");
        String gridIncludeHeader = getPropertyString("gridIncludeHeader");
        String gridDirection = getPropertyString("gridDirection");
        String wordFileName = getPropertyString("wordFileName");
        String gridWidth = getPropertyString("gridWidth");
        String imageWidth = getPropertyString("imageWidth");
        String imageHeight = getPropertyString("imageHeight");
        String pathFormDefId = getPropertyString("pathFormDefId");

        if (wordFileName == "" || wordFileName.trim().isEmpty()) {
            wordFileName = "Doc File";
        }

        if ("FILE_PATH".equalsIgnoreCase(pathOptions)) {
            DocumentGenerationUtil.generateSingleFile(request, response, recordId,
                    formDefId, templateFile,
                    gridIncludeHeader, gridDirection,
                    filePath + "/" + wordFileName, gridWidth,
                    imageWidth, imageHeight, "file");
            File outputFile = DocumentGenerationUtil.getGeneratedFile();
            if (outputFile.exists()) {
                LogUtil.info(getClassName(), "File saved to: " + filePath);
            } 

        } else if ("FORM_FIELD".equalsIgnoreCase(pathOptions)) {
            DocumentGenerationUtil.generateSingleFile(request, response, recordId,
                    formDefId, templateFile,
                    gridIncludeHeader, gridDirection,
                    wordFileName, gridWidth,
                    imageWidth, imageHeight, "file");

            File outputFile = DocumentGenerationUtil.getGeneratedFile();

            if (outputFile.exists()) {
                String fileName = outputFile.getName();
                String tableName = appService.getFormTableName(appDef, pathFormDefId);
                FileUtil.storeFile(outputFile, tableName, recordId);
                FormRowSet rows = new FormRowSet();
                FormRow row = new FormRow();
                row.setId(recordId);
                row.put(fileFieldId, fileName);
                rows.add(row);
                appService.storeFormData(pathFormDefId, tableName, rows, recordId);
                LogUtil.info(getClassName(), "File saved to form");
            }
        }
        return null;
    }

    @Override
    public String getName() {
        return AppPluginUtil.getMessage("org.joget.marketplace.DocumentGenerationTool.pluginLabel", getClassName(), MESSAGE_PATH);
    }

    @Override
    public String getVersion() {
        return Activator.VERSION;
    }

    @Override
    public String getDescription() {
        return AppPluginUtil.getMessage("org.joget.marketplace.DocumentGenerationTool.pluginDesc", getClassName(), MESSAGE_PATH);
    }

    @Override
    public String getLabel() {
        return AppPluginUtil.getMessage("org.joget.marketplace.DocumentGenerationTool.pluginLabel", getClassName(), MESSAGE_PATH);
    }

    @Override
    public String getClassName() {
        return this.getClass().getName();
    }

    @Override
    public String getPropertyOptions() {
        return AppUtil.readPluginResource(getClass().getName(), "/properties/documentGenerationTool.json", null, true, MESSAGE_PATH);
    }

    private static String getBaseName(String filename) {
        int dotIndex = filename.lastIndexOf('.');
        return (dotIndex == -1) ? filename : filename.substring(0, dotIndex);
    }

    private static String getExtension(String filename) {
        int dotIndex = filename.lastIndexOf('.');
        return (dotIndex == -1) ? "" : filename.substring(dotIndex);
    }
}

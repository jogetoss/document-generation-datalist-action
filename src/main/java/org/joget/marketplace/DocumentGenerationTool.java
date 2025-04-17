package org.joget.marketplace;

import java.io.File;
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
import org.joget.commons.util.UuidGenerator;
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
        if (recordId == null && recordId.equals("")) {
            if (wfAssignment != null) {
                recordId = appService.getOriginProcessId(wfAssignment.getProcessId());
            }
        }

        DocumentGenerationUtil.generateSingleFile(request, response, recordId,
                getPropertyString("formDefId"), getPropertyString("templateFile"),
                getPropertyString("gridIncludeHeader"), getPropertyString("gridDirection"),
                getPropertyString("wordFileName"), getPropertyString("gridWidth"),
                getPropertyString("imageWidth"), getPropertyString("imageHeight"), "file");

        // file output
        File outputFile = DocumentGenerationUtil.getGeneratedFile();
        String filePath = getPropertyString("filePath");;
        String formDefId = getPropertyString("formDefId");
        String fileFieldId = getPropertyString("fileFieldId");

        if (outputFile.exists()) {
            String fileName = outputFile.getName();
            String tableName = appService.getFormTableName(appDef, formDefId);
            //String id = UuidGenerator.getInstance().getUuid();
            FileUtil.storeFile(outputFile, tableName, recordId);
            FormRowSet rows = new FormRowSet();
            FormRow row = new FormRow();
            row.put(fileFieldId, fileName);
            rows.add(row);
            appService.storeFormData(formDefId, tableName, rows, recordId);
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
}

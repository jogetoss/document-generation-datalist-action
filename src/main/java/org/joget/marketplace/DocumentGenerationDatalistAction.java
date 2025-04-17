package org.joget.marketplace;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import org.joget.apps.app.service.AppPluginUtil;
import org.joget.apps.app.service.AppUtil;
import org.joget.apps.datalist.model.DataList;
import org.joget.apps.datalist.model.DataListActionDefault;
import org.joget.apps.datalist.model.DataListActionResult;
import org.joget.commons.util.LogUtil;
import org.joget.workflow.util.WorkflowUtil;

import org.joget.marketplace.util.DocumentGenerationUtil;

public class DocumentGenerationDatalistAction extends DataListActionDefault {

    private final static String MESSAGE_PATH = "message/DocumentGenerationDatalistAction";

    @Override
    public String getName() {
        return AppPluginUtil.getMessage("org.joget.marketplace.DocumentGenerationDatalistAction.pluginLabel", getClassName(), MESSAGE_PATH);
    }

    @Override
    public String getVersion() {
        return Activator.VERSION;
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
        String recordIdColumn = getPropertyString("recordIdColumn");
        if ("id".equalsIgnoreCase(recordIdColumn) || recordIdColumn.isEmpty()) {
            return getPropertyString("hrefColumn");
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
        return AppUtil.readPluginResource(getClassName(), "/properties/documentGenerationDatalistAction.json", null, true, MESSAGE_PATH);
    }

    @Override
    public DataListActionResult executeAction(DataList dataList, String[] rowKeys) {

        HttpServletRequest request = WorkflowUtil.getHttpServletRequest();
        if (request != null && !"POST".equalsIgnoreCase(request.getMethod())) {
            return null;
        }

        if (rowKeys != null && rowKeys.length > 0) {
            try {
                HttpServletResponse response = WorkflowUtil.getHttpServletResponse();

                if (rowKeys.length == 1) {
                    DocumentGenerationUtil.generateSingleFile(request, response, rowKeys[0],
                            getPropertyString("formDefId"), getPropertyString("templateFile"),
                            getPropertyString("gridIncludeHeader"), getPropertyString("gridDirection"),
                            getPropertyString("wordFileName"), getPropertyString("gridWidth"),
                            getPropertyString("imageWidth"), getPropertyString("imageHeight"), "");
                } else {
                    DocumentGenerationUtil.generateMultipleFile(request, response, rowKeys,
                            getPropertyString("formDefId"), getPropertyString("templateFile"),
                            getPropertyString("gridIncludeHeader"), getPropertyString("gridDirection"),
                            getPropertyString("zipFileName"), getPropertyString("gridWidth"),
                            getPropertyString("imageWidth"), getPropertyString("imageHeight"));
                }
            } catch (Exception e) {
                LogUtil.error(getClassName(), e, "Failed to generate word file");
            }
        }
        return null;
    }

   
}

package com.poi.excel.util.jxls;

import org.apache.poi.ss.usermodel.Workbook;
import org.jxls.common.Context;
import org.jxls.util.JxlsHelper;
import org.springframework.core.io.ClassPathResource;
import org.springframework.http.HttpHeaders;
import org.springframework.web.servlet.view.document.AbstractXlsxView;
import tech.simter.jxls.ext.CommonFunctions;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.BufferedInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.Map;

public class JxlsExcelDownload extends AbstractXlsxView {

    @Override
    protected void buildExcelDocument(Map<String, Object> model, Workbook workbook, HttpServletRequest request, HttpServletResponse response) throws Exception {
        String templateName = (String) model.get("template");
        ClassPathResource classPathResource = new ClassPathResource(String.format("static/excel/templates/%s.xlsx", templateName));

        try (InputStream inputStream = new BufferedInputStream(classPathResource.getInputStream())) {
            response.setHeader(HttpHeaders.CONTENT_DISPOSITION, String.format("attachment;filename=%s.xlsx", templateName));
            OutputStream outputStream = response.getOutputStream();

            Context context = new Context();
            context.putVar("fn", CommonFunctions.getSingleton());
            context.putVar("list", model.get("list"));
            JxlsHelper.getInstance().processTemplate(inputStream, outputStream, context);
            outputStream.close();
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }
}
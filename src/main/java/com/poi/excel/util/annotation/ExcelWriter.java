package com.poi.excel.util.annotation;

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.servlet.http.HttpServletResponse;
import java.io.Serializable;
import java.io.UnsupportedEncodingException;
import java.lang.reflect.Field;
import java.text.DateFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.*;

@Slf4j
public class ExcelWriter {

    private final Workbook workbook;
    private final Map<String, Object> data;
    private final HttpServletResponse response;

    public ExcelWriter(Workbook workbook, Map<String, Object> data, HttpServletResponse response) {
        this.workbook = workbook;
        this.data = data;
        this.response = response;
    }

    // 엑셀 파일 생성
    public void create() throws UnsupportedEncodingException {
        setFileName(response, mapToFileName());

        Sheet sheet = workbook.createSheet();

        createHead(sheet, mapToHeadList());

        createBody(sheet, mapToBodyList());
    }

    // 모델 객체에서 파일 이름 꺼내기
    private String mapToFileName() {
        return (String) data.get("filename");
    }

    // 모델 객체에서 헤더 이름 리스트 꺼내기
    private List<String> mapToHeadList() {
        return (List<String>) data.get("head");
    }

    // 모델 객체에서 바디 데이터 리스트 꺼내기
    private List<List<String>> mapToBodyList() {
        return (List<List<String>>) data.get("body");
    }

    // 파일 이름 지정
    private void setFileName(HttpServletResponse response, String fileName) throws UnsupportedEncodingException {
        fileName += "_" + new SimpleDateFormat("yyyyMMddHHmmss").format(new Date());
        fileName = new String(getFileExtension(fileName).getBytes("KSC5601"), "8859_1");
        response.setHeader("Content-Disposition",
                "attachment; filename=\"" + fileName + "\"");
    }

    // 넘어온 뷰에 따라서 확장자 결정
    private String getFileExtension(String fileName) {
        if (workbook instanceof XSSFWorkbook) {
            fileName += ".xlsx";
        }
        if (workbook instanceof SXSSFWorkbook) {
            fileName += ".xlsx";
        }
        if (workbook instanceof HSSFWorkbook) {
            fileName += ".xls";
        }

        return fileName;
    }

    // 엑셀 헤더 생성
    private void createHead(Sheet sheet, List<String> headList) {
        createRow(sheet, headList, 0);
    }

    // 엑셀 바디 생성
    private void createBody(Sheet sheet, List<List<String>> bodyList) {
        int rowSize = bodyList.size();
        for (int i = 0; i < rowSize; i++) {
            createRow(sheet, bodyList.get(i), i + 1);
        }
    }

    // 행 생성
    private void createRow(Sheet sheet, List<String> cellList, int rowNum) {
        int size = cellList.size();
        Row row = sheet.createRow(rowNum);

        for (int i = 0; i < size; i++) {
            Cell cell = row.createCell(i);
            String value = cellList.get(i);
            if (isValidDate(value)) {
                try {
                    DateFormat sdf1 = new SimpleDateFormat("yyyyMMddHHmmss", Locale.KOREA);
                    DateFormat sdf2 = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss", Locale.KOREA);
                    cell.setCellValue(sdf2.format(sdf1.parse(value)));
                } catch (ParseException pe) {
                    log.debug("Date Parse Error : " + value);
                }
            } else if (isNumeric(value)) {
                CellStyle cellStyle = workbook.createCellStyle();
                DataFormat df = workbook.createDataFormat();
                cellStyle.setDataFormat(df.getFormat("#,##0"));
                cell.setCellStyle(cellStyle);
                cell.setCellValue(new Double(value));
            } else {
                cell.setCellValue(value);
            }
        }
    }

    // 모델 객체에 담을 형태로 엑셀 데이터 생성
    public static Map<String, Object> createExcelData(List<? extends ExcelDto> data, Class<?> target) {
        Map<String, Object> excelData = new HashMap<>();
        excelData.put("filename", createFileName(target));
        excelData.put("head", createHeaderName(target));
        excelData.put("body", createBodyData(data));
        return excelData;
    }

    // @ExcelColumnName에서 헤더 이름 리스트 생성
    private static List<String> createHeaderName(Class<?> header) {
        List<String> headData = new ArrayList<>();
        for (Field field : header.getDeclaredFields()) {
            field.setAccessible(true);
            if (field.isAnnotationPresent(ExcelColumnName.class)) {
                String headerName = field.getAnnotation(ExcelColumnName.class).headerName();
                if (headerName.equals("")) {
                    headData.add(field.getName());
                } else {
                    headData.add(headerName);
                }
            }
        }
        return headData;
    }

    // @ExcelFileName 에서 엑셀 파일 이름 생성
    private static String createFileName(Class<?> file) {
        if (file.isAnnotationPresent(ExcelFileName.class)) {
            String filename = file.getAnnotation(ExcelFileName.class).filename();
            return filename.equals("") ? file.getSimpleName() : filename;
        }
        throw new RuntimeException("excel filename not exist");
    }

    // 데이터 리스트 형태로 가공
    private static List<List<Serializable>> createBodyData(List<? extends ExcelDto> dataList) {
        List<List<Serializable>> bodyData = new ArrayList<>();
        dataList.forEach(v -> bodyData.add(v.mapToList()));
        return bodyData;
    }

    public boolean isNumeric(String strNum) {
        if (strNum == null) {
            return false;
        }
        try {
            double d = Double.parseDouble(strNum);
        } catch (NumberFormatException nfe) {
            return false;
        }
        return true;
    }

    public boolean isValidDate(String dateStr) {
        if (null == dateStr) {
            return false;
        }

        DateFormat sdf = new SimpleDateFormat("yyyyMMddHHmmss", Locale.KOREA);
        sdf.setLenient(false);
        try {
            sdf.parse(dateStr);
        } catch (ParseException e) {
            return false;
        }
        return true;
    }
}
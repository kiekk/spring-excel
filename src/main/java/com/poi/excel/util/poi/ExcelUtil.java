package com.poi.excel.util.poi;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import javax.servlet.http.HttpServletResponse;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;
import java.util.Map;

public class ExcelUtil {

    private int rowNum = 0;

    public void createExcelToFile(List<Map<String, Object>> dataList, String filepath) throws FileNotFoundException, IOException {

        Workbook workbook = new SXSSFWorkbook();
        Sheet sheet = workbook.createSheet("엑셀_다운로드_POI");

        rowNum = 0;

        createExcel(sheet, dataList);

        FileOutputStream fos = new FileOutputStream(filepath);
        workbook.write(fos);
        workbook.close();

    }

    public void createExcelToResponse(List<Map<String, Object>> dataList, String filename, HttpServletResponse response) throws IOException {

        Workbook workbook = new SXSSFWorkbook();
        Sheet sheet = workbook.createSheet("엑셀_다운로드_POI");

        rowNum = 0;

        createExcel(sheet, dataList);

        response.setContentType("application/vnd.ms-excel");
        response.setHeader("Content-Disposition", String.format("attachment;filename=%s.xlsx", filename));

        workbook.write(response.getOutputStream());
        workbook.close();
    }

    private void createExcel(Sheet sheet, List<Map<String, Object>> dataList) {

        for (Map<String, Object> data : dataList) {
            Row row = sheet.createRow(rowNum++);
            int cellNum = 0;

            for (String key : data.keySet()) {
                Cell cell = row.createCell(cellNum++);

                cell.setCellValue(data.get(key).toString());
            }
        }
    }

}

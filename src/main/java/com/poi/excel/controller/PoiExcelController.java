package com.poi.excel.controller;

import com.poi.excel.entity.Board;
import com.poi.excel.util.poi.ExcelUtil;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.ZoneId;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

@Controller
@RequestMapping("/poi")
public class PoiExcelController {

    private final static List<Map<String, Object>> headerList1 = new ArrayList<>();
    private final static List<Map<String, Object>> boardList1 = new ArrayList<>();
    private final static List<Board> boardList2 = new ArrayList<>();

    static {
        for (int i = 1; i < 10; i++) {
            boardList1.add(
                    Board.builder().id(i).title("게시글_" + i).content("게시글내용_" + i).writer("작성자_" + i)
                            .viewCount(i).likeIt(i).createDate(LocalDateTime.now()).updateDate(LocalDateTime.now()).build().entityToMap());
            boardList2.add(
                    Board.builder().id(i).title("게시글_" + i).content("게시글내용_" + i).writer("작성자_" + i)
                            .viewCount(i).likeIt(i).createDate(LocalDateTime.now()).updateDate(LocalDateTime.now()).build());
        }
        headerList1.add(new Board().getHeaderToMap());
    }

    @GetMapping("excel-upload")
    public String poiExcelUpload() {
        return "poi-excel-upload";
    }

    @PostMapping("excel-upload")
    public void poiExcelUpload(MultipartFile excelFile) {
        List<Board> boardList = new ArrayList<>();
        try {
            OPCPackage opcPackage = OPCPackage.open(excelFile.getInputStream());
            XSSFWorkbook workbook = new XSSFWorkbook(opcPackage);

            // 첫 번째 시트 불러오기
            XSSFSheet sheet = workbook.getSheetAt(0);

            for (int i = 1; i < sheet.getLastRowNum() + 1; i++) {
                Board board = new Board();
                XSSFRow row = sheet.getRow(i);
                int rowNum = 0;

                if (null == row) {
                    continue;
                }

                // 행의 1번째 열
                XSSFCell cell = row.getCell(rowNum++);
                if (null != cell) {
                    board.setId(Integer.parseInt(cell.getStringCellValue()));
                }

                cell = row.getCell(rowNum++);
                if (null != cell) {
                    board.setTitle(cell.getStringCellValue());
                }

                cell = row.getCell(rowNum++);
                if (null != cell) {
                    board.setTitle(cell.getStringCellValue());
                }

                cell = row.getCell(rowNum++);
                if (null != cell) {
                    board.setContent(cell.getStringCellValue());
                }

                cell = row.getCell(rowNum++);
                if (null != cell) {
                    board.setWriter(cell.getStringCellValue());
                }

                cell = row.getCell(rowNum++);
                if (null != cell) {
                    board.setViewCount(Integer.parseInt(cell.getStringCellValue()));
                }

                cell = row.getCell(rowNum++);
                if (null != cell) {
                    board.setLikeIt(Integer.parseInt(cell.getStringCellValue()));
                }

                cell = row.getCell(rowNum++);
                if (null != cell) {
                    board.setCreateDate(cell.getDateCellValue().toInstant().atZone(ZoneId.systemDefault()).toLocalDateTime());
                }

                cell = row.getCell(rowNum++);
                if (null != cell) {
                    board.setUpdateDate(cell.getDateCellValue().toInstant().atZone(ZoneId.systemDefault()).toLocalDateTime());
                }

                boardList.add(board);
            }
        } catch (InvalidFormatException | IOException e) {
            e.printStackTrace();
        }

        System.out.println("################################################");
        System.out.println("################################################");
        System.out.println("################################################");
        System.out.println("################################################");
        System.out.println("################################################");
        boardList.forEach(System.out::println);
    }

    @RequestMapping("excel-download-1")
    public void poiExcelDownload1(HttpServletResponse response) throws IOException {

        new ExcelUtil().createExcelToResponse(
                headerList1,
                boardList1,
                String.format("%s-%s", "data", LocalDate.now().toString()),
                response
        );
    }

    @RequestMapping("excel-create-file")
    public String poiExcelCreateFile() throws IOException {
        String filepath = "C:/Users/soonho/Desktop/data.xlsx";
        new ExcelUtil().createExcelToFile(headerList1, boardList1, filepath);

        return "redirect:/";
    }

    @RequestMapping("excel-download-2")
    public void poiExcelDownload2(HttpServletResponse response) throws IOException {
        final String fileName = "boardList.xlsx";

        /* 엑셀 그리기 */
        final String[] colNames = {
                "아이디", "제목", "내용", "작성자", "조회수", "좋아요수", "등록일", "생성일"
        };

        // 헤더 사이즈
        final int[] colWidths = {
                3000, 5000, 5000, 3000, 3000, 5000, 5000, 5000
        };

        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet;
        XSSFCell cell;
        XSSFRow row;

        //Font
        Font fontHeader = workbook.createFont();
        fontHeader.setFontName("맑은 고딕");
        fontHeader.setFontHeight((short) (9 * 20));
        fontHeader.setBold(true);
        Font font9 = workbook.createFont();
        font9.setFontName("맑은 고딕");
        font9.setFontHeight((short) (9 * 20));

        // 엑셀 헤더 셋팅
        XSSFCellStyle headerStyle = workbook.createCellStyle();
        headerStyle.setAlignment(HorizontalAlignment.CENTER);
        headerStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        headerStyle.setBorderRight(BorderStyle.THIN);
        headerStyle.setBorderLeft(BorderStyle.THIN);
        headerStyle.setBorderTop(BorderStyle.THIN);
        headerStyle.setBorderBottom(BorderStyle.THIN);
        headerStyle.setFillForegroundColor(new XSSFColor(new java.awt.Color(226, 239, 218), new DefaultIndexedColorMap()));
        headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        headerStyle.setFont(fontHeader);

        // 엑셀 바디 셋팅
        CellStyle bodyStyle = workbook.createCellStyle();
        bodyStyle.setAlignment(HorizontalAlignment.CENTER);
        bodyStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        bodyStyle.setBorderRight(BorderStyle.THIN);
        bodyStyle.setBorderLeft(BorderStyle.THIN);
        bodyStyle.setBorderTop(BorderStyle.THIN);
        bodyStyle.setBorderBottom(BorderStyle.THIN);
        bodyStyle.setFont(font9);

        // 엑셀 왼쪽 설정
        CellStyle leftStyle = workbook.createCellStyle();
        leftStyle.setAlignment(HorizontalAlignment.LEFT);
        leftStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        leftStyle.setBorderRight(BorderStyle.THIN);
        leftStyle.setBorderLeft(BorderStyle.THIN);
        leftStyle.setBorderTop(BorderStyle.THIN);
        leftStyle.setBorderBottom(BorderStyle.THIN);
        leftStyle.setFont(font9);

        //rows
        int rowCnt = 0;
        int cellCnt = 0;
        int listCount = boardList2.size();

        // 엑셀 시트명 설정
        sheet = workbook.createSheet("사용자현황");
        row = sheet.createRow(rowCnt++);
        //헤더 정보 구성
        for (int i = 0; i < colNames.length; i++) {
            cell = row.createCell(i);
            cell.setCellStyle(headerStyle);
            cell.setCellValue(colNames[i]);
            sheet.setColumnWidth(i, colWidths[i]);
        }
        //데이터 부분 생성
        for (Board board : boardList2) {
            cellCnt = 0;
            row = sheet.createRow(rowCnt++);
            // id
            cell = row.createCell(cellCnt++);
            cell.setCellStyle(bodyStyle);
            cell.setCellValue(board.getId());

            // 제목
            cell = row.createCell(cellCnt++);
            cell.setCellStyle(bodyStyle);
            cell.setCellValue(board.getTitle());

            // 내용
            cell = row.createCell(cellCnt++);
            cell.setCellStyle(bodyStyle);
            cell.setCellValue(board.getContent());

            // 작성자
            cell = row.createCell(cellCnt++);
            cell.setCellStyle(bodyStyle);
            cell.setCellValue(board.getWriter());

            // 조회수
            cell = row.createCell(cellCnt++);
            cell.setCellStyle(bodyStyle);
            cell.setCellValue(board.getViewCount());

            // 좋아요수
            cell = row.createCell(cellCnt++);
            cell.setCellStyle(bodyStyle);
            cell.setCellValue(board.getLikeIt());

            // 등록일
            cell = row.createCell(cellCnt++);
            cell.setCellStyle(bodyStyle);
            cell.setCellValue((board.getCreateDate()).toString());

            // 수정일
            cell = row.createCell(cellCnt++);
            cell.setCellStyle(bodyStyle);
            cell.setCellValue((board.getCreateDate()).toString());
        }

        response.setContentType("application/vnd.ms-excel");
        // 엑셀 파일명 설정
        response.setHeader("Content-Disposition", "attachment;filename=" + fileName);
        try {
            workbook.write(response.getOutputStream());
        } catch (IOException e) {
            e.printStackTrace();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

}

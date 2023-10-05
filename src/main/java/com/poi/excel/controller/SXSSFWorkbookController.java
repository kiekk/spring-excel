package com.poi.excel.controller;

import com.poi.excel.entity.Board;
import com.poi.excel.entity.User;
import com.poi.excel.util.annotation.ExcelDownload;
import com.poi.excel.util.annotation.ExcelDownloadNew;
import com.poi.excel.util.annotation.ExcelWriter;
import com.poi.excel.util.annotation.ExcelWriterNew;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.servlet.ModelAndView;

import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

@RequestMapping("sxssf")
@RestController
public class SXSSFWorkbookController {

    @RequestMapping("excel-download-user")
    public ModelAndView sxssfExcelDownloadUser() {
        List<User> userList = new ArrayList<>();

        for (int i = 1; i < 1_040_000; i++) {
            userList.add(
                    User.builder().id(i + "").loginId("user_" + i).name("user_name_" + i).address("사용자 주소_" + i).mobile("010-0000-000" + i)
                            .createDate(LocalDateTime.now()).updateDate(LocalDateTime.now()).build());
        }
        Map<String, Object> excelData = ExcelWriterNew.createExcelData(userList, User.class);
        return new ModelAndView(new ExcelDownloadNew(), excelData);
    }

    @RequestMapping("excel-download-board")
    public ModelAndView sxssfExcelDownloadBoard() {
        List<Board> boardList = new ArrayList<>();
        for (int i = 1; i < 1_040_000; i++) {
            boardList.add(
                    Board.builder().id(i).title("게시글_" + i).content("게시글내용_" + i).writer("작성자_" + i)
                            .viewCount(i).likeIt(i).createDate(LocalDateTime.now()).updateDate(LocalDateTime.now()).build());
        }
        Map<String, Object> excelData = ExcelWriterNew.createExcelData(boardList, Board.class);
        return new ModelAndView(new ExcelDownloadNew(), excelData);
    }

}

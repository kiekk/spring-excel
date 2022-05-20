package com.poi.excel.controller;

import com.poi.excel.entity.Board;
import com.poi.excel.entity.User;
import com.poi.excel.util.annotation.ExcelDownload;
import com.poi.excel.util.annotation.ExcelWriter;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.servlet.ModelAndView;

import java.io.IOException;
import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

@Controller
@RequestMapping("/annotation")
public class AnnotationController {

    private final static List<Board> boardList = new ArrayList<>();
    private final static List<User> userList = new ArrayList<>();

    static {
        for (int i = 1; i < 10; i++) {
            boardList.add(
                    Board.builder().id(i).title("게시글_" + i).content("게시글내용_" + i).writer("작성자_" + i)
                            .viewCount(i).likeIt(i).createDate(LocalDateTime.now()).updateDate(LocalDateTime.now()).build());
            userList.add(
                    User.builder().id(i + "").loginId("user_" + i).name("user_name_" + i).address("사용자 주소_" + i).mobile("010-0000-000" + i)
                            .createDate(LocalDateTime.now()).updateDate(LocalDateTime.now()).build());
        }
    }

    @RequestMapping("excel-download-board")
    public ModelAndView poiExcelDownloadBoard() throws IOException {
        Map<String, Object> excelData = ExcelWriter.createExcelData(boardList, Board.class);
        return new ModelAndView(new ExcelDownload(), excelData);
    }

    @RequestMapping("excel-download-user")
    public ModelAndView poiExcelDownloadUser() throws IOException {
        Map<String, Object> excelData = ExcelWriter.createExcelData(userList, User.class);
        return new ModelAndView(new ExcelDownload(), excelData);
    }
}

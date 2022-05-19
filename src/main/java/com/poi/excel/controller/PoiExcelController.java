package com.poi.excel.controller;

import com.poi.excel.entity.Board;
import com.poi.excel.util.poi.ExcelUtil;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

@Controller
@RequestMapping("/poi")
public class PoiExcelController {

    private final static List<Map<String, Object>> headerList = new ArrayList<>();
    private final static List<Map<String, Object>> boardList = new ArrayList<>();

    static {
        boardList.add(
                Board.builder().id(1).title("게시글1").content("게시글내용1").writer("작성자1")
                        .viewCount(0).likeIt(1).createDate(LocalDateTime.now()).updateDate(LocalDateTime.now()).build().entityToMap());
        boardList.add(
                Board.builder().id(2).title("게시글2").content("게시글내용2").writer("작성자2")
                        .viewCount(1).likeIt(2).createDate(LocalDateTime.now()).updateDate(LocalDateTime.now()).build().entityToMap());
        boardList.add(
                Board.builder().id(3).title("게시글3").content("게시글내용3").writer("작성자3")
                        .viewCount(2).likeIt(3).createDate(LocalDateTime.now()).updateDate(LocalDateTime.now()).build().entityToMap());
        boardList.add(
                Board.builder().id(4).title("게시글4").content("게시글내용4").writer("작성자4")
                        .viewCount(3).likeIt(4).createDate(LocalDateTime.now()).updateDate(LocalDateTime.now()).build().entityToMap());
        boardList.add(
                Board.builder().id(5).title("게시글5").content("게시글내용5").writer("작성자5")
                        .viewCount(4).likeIt(5).createDate(LocalDateTime.now()).updateDate(LocalDateTime.now()).build().entityToMap());
        boardList.add(
                Board.builder().id(6).title("게시글6").content("게시글내용6").writer("작성자6")
                        .viewCount(5).likeIt(6).createDate(LocalDateTime.now()).updateDate(LocalDateTime.now()).build().entityToMap());
        boardList.add(
                Board.builder().id(7).title("게시글7").content("게시글내용7").writer("작성자7")
                        .viewCount(6).likeIt(7).createDate(LocalDateTime.now()).updateDate(LocalDateTime.now()).build().entityToMap());
        boardList.add(
                Board.builder().id(8).title("게시글8").content("게시글내용8").writer("작성자8")
                        .viewCount(7).likeIt(8).createDate(LocalDateTime.now()).updateDate(LocalDateTime.now()).build().entityToMap());
        boardList.add(
                Board.builder().id(9).title("게시글9").content("게시글내용9").writer("작성자9")
                        .viewCount(8).likeIt(9).createDate(LocalDateTime.now()).updateDate(LocalDateTime.now()).build().entityToMap());
        boardList.add(
                Board.builder().id(10).title("게시글10").content("게시글내용10").writer("작성자10")
                        .viewCount(9).likeIt(10).createDate(LocalDateTime.now()).updateDate(LocalDateTime.now()).build().entityToMap());

        headerList.add(new Board().getHeaderToMap());
    }

    @RequestMapping("excel-download")
    public void poiExcelDownload(HttpServletResponse response) throws IOException {

        new ExcelUtil().createExcelToResponse(
                headerList,
                boardList,
                String.format("%s-%s", "data", LocalDate.now().toString()),
                response
        );
    }

    @RequestMapping("excel-create-file")
    public String poiExcelCreateFile() throws IOException {
        String filepath = "C:/Users/soonho/Desktop/data.xlsx";
        new ExcelUtil().createExcelToFile(headerList, boardList, filepath);

        return "redirect:/";
    }

}

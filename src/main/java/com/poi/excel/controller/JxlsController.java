package com.poi.excel.controller;

import com.poi.excel.entity.Board;
import com.poi.excel.entity.User;
import com.poi.excel.util.LocalDateTimeConverter;
import com.poi.excel.util.jxls.JxlsExcelDownload;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.jxls.reader.ReaderBuilder;
import org.jxls.reader.XLSReader;
import org.springframework.core.io.ClassPathResource;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.ResponseBody;
import org.springframework.web.multipart.MultipartFile;
import org.springframework.web.servlet.ModelAndView;
import org.xml.sax.SAXException;

import java.io.BufferedInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

@Controller
@RequestMapping("/jxls")
public class JxlsController {

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
    public ModelAndView jxlsExcelDownloadBoard(Model model) {
        model.addAttribute("list", boardList);
        model.addAttribute("template", "board");

        return new ModelAndView(new JxlsExcelDownload());
    }

    @RequestMapping("excel-download-user")
    public ModelAndView jxlsExcelDownloadUser(Model model) {
        model.addAttribute("list", userList);
        model.addAttribute("template", "user");

        return new ModelAndView(new JxlsExcelDownload());
    }

    @GetMapping("excel-upload")
    public String jxlsExcelUpload() {
        return "jxls-excel-upload";
    }

    @PostMapping("excel-upload")
    @ResponseBody
    public void jxlsExcelUpload(MultipartFile excelFile) {
        ClassPathResource classPathResource = new ClassPathResource("static/excel/xml/board.xml");
        List<Board> uploadBoardList = new ArrayList<>();

        try (InputStream inputStream = new BufferedInputStream(classPathResource.getInputStream());
             InputStream multipartInputStream = excelFile.getInputStream()) {
            XLSReader xlsReader = ReaderBuilder.buildFromXML(inputStream);
            xlsReader.getConvertUtilsBeanProvider().getConvertUtilsBean().register(new LocalDateTimeConverter(), java.time.LocalDateTime.class);

            Map<String, Object> beans = new HashMap<>();
            beans.put("boardList", uploadBoardList);

            xlsReader.read(multipartInputStream, beans);

            System.out.println("##########################");
            System.out.println("##########################");
            System.out.println("##########################");
            System.out.println("##########################");
            uploadBoardList.forEach(System.out::println);
        } catch (IOException | SAXException | InvalidFormatException e) {
            throw new RuntimeException(e);
        }
    }

}


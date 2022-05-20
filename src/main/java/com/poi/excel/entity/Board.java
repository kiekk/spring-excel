package com.poi.excel.entity;

import com.poi.excel.util.annotation.ExcelColumnName;
import com.poi.excel.util.annotation.ExcelDto;
import com.poi.excel.util.annotation.ExcelFileName;
import lombok.*;

import java.io.Serializable;
import java.time.LocalDateTime;
import java.util.Arrays;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

@Getter
@Setter
@ToString
@NoArgsConstructor
@AllArgsConstructor
@Builder
@ExcelFileName(filename = "게시판목록")
public class Board implements ExcelDto {

    @ExcelColumnName(headerName = "아이디")
    private int id;
    @ExcelColumnName(headerName = "제목")
    private String title;
    @ExcelColumnName(headerName = "내용")
    private String content;
    @ExcelColumnName(headerName = "작성자")
    private String writer;
    @ExcelColumnName(headerName = "조회수")
    private int viewCount;
    @ExcelColumnName(headerName = "좋아요수")
    private int likeIt;
    @ExcelColumnName(headerName = "등록일")
    private LocalDateTime createDate;
    @ExcelColumnName(headerName = "수정일")
    private LocalDateTime updateDate;

    public Map<String, Object> entityToMap() {
        Map<String, Object> result = new HashMap<>();
        result.put("id", id);
        result.put("title", title);
        result.put("content", content);
        result.put("writer", writer);
        result.put("viewCount", viewCount);
        result.put("likeIt", likeIt);
        result.put("createDate", createDate);
        result.put("updateDate", updateDate);
        return result;
    }

    public Map<String, Object> getHeaderToMap() {
        Map<String, Object> result = new HashMap<>();
        result.put("id", "아이디");
        result.put("title", "제목");
        result.put("content", "내용");
        result.put("writer", "작성자");
        result.put("viewCount", "조회수");
        result.put("likeIt", "좋아요수");
        result.put("createDate", "등록일");
        result.put("updateDate", "수정일");
        return result;
    }

    @Override
    public List<Serializable> mapToList() {
        return Arrays.asList(
                String.valueOf(id),
                title,
                content,
                writer,
                String.valueOf(viewCount),
                String.valueOf(likeIt),
                createDate.toString(),
                updateDate.toString()
        );
    }
}

package com.poi.excel.entity;

import lombok.*;

import java.time.LocalDateTime;
import java.util.HashMap;
import java.util.Map;

@Getter
@Setter
@ToString
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class Board {

    private int id;
    private String title;
    private String content;
    private String writer;
    private int viewCount;
    private int likeIt;
    private LocalDateTime createDate;
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

}

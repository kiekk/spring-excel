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

}

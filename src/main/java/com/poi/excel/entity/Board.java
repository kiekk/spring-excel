package com.poi.excel.entity;

import lombok.*;

import java.time.LocalDateTime;

@Getter
@Setter
@ToString
@NoArgsConstructor
@AllArgsConstructor
public class Board {

    private String id;
    private String title;
    private String content;
    private String writer;
    private int viewCount;
    private int likeIt;
    private LocalDateTime createDate;
    private LocalDateTime updateDate;

}

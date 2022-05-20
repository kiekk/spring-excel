package com.poi.excel.entity;

import com.poi.excel.util.annotation.ExcelColumnName;
import com.poi.excel.util.annotation.ExcelDto;
import com.poi.excel.util.annotation.ExcelFileName;
import lombok.*;

import java.io.Serializable;
import java.time.LocalDateTime;
import java.util.Arrays;
import java.util.List;

@Getter
@Setter
@ToString
@NoArgsConstructor
@AllArgsConstructor
@Builder
@ExcelFileName(filename = "유저 목록")
public class User implements ExcelDto {

    @ExcelColumnName(headerName = "아이디")
    private String id;
    @ExcelColumnName(headerName = "로그인 아이디")
    private String loginId;
    @ExcelColumnName(headerName = "이름")
    private String name;
    @ExcelColumnName(headerName = "전화번호")
    private String mobile;
    @ExcelColumnName(headerName = "주소")
    private String address;
    @ExcelColumnName(headerName = "등록일")
    private LocalDateTime createDate;
    @ExcelColumnName(headerName = "수정일")
    private LocalDateTime updateDate;

    @Override
    public List<Serializable> mapToList() {
        return Arrays.asList(
                id,
                loginId,
                name,
                mobile,
                address,
                createDate.toString(),
                updateDate.toString()
        );
    }
}

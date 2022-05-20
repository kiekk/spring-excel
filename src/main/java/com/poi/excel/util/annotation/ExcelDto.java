package com.poi.excel.util.annotation;

import java.io.Serializable;
import java.util.List;

public interface ExcelDto {
    List<Serializable> mapToList();
}

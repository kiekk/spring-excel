package com.poi.excel.util;

import org.apache.commons.beanutils.ConversionException;
import org.apache.commons.beanutils.Converter;

import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;

public class LocalDateTimeConverter implements Converter {

    public Object convert(Class type, Object value) {
        if( value == null ) {
            throw new ConversionException("No value specified");
        }
        try {
            String str = (String) value;
            DateTimeFormatter formatter = DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss");
            return LocalDateTime.parse(str, formatter);
        } catch (Exception e) {
            e.printStackTrace();
        }
        return null;
    }
}

package com.gstsgy.tools.excel.bean;


import com.gstsgy.tools.excel.function.GetFunction;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;

public class DateNow<T> implements GetFunction<T,Object> {
    @Override
    public Object getVal(T t) {
        return LocalDate.now().format(DateTimeFormatter.ofPattern("yyyy-MM-dd"));
    }
}

package com.github.excel.help.bean;


import com.github.excel.help.function.GetFunction;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;

public class DateTimeNow<T> implements GetFunction<T,Object> {
    @Override
    public Object getVal(T t) {
        return  LocalDateTime.now().format(DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss"));
    }
}

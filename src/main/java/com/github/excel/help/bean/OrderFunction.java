package com.github.excel.help.bean;

import com.github.excel.help.function.GetFunction;

public class OrderFunction<T> implements GetFunction<T, Object> {
    private  int start=1;
    private int sep;
    public OrderFunction() {
    }

    public OrderFunction(int start, int sep) {
        this.start = start;
        this.sep = sep;
    }

    @Override
    public Object getVal(T o) {
        return start += sep;
    }
}

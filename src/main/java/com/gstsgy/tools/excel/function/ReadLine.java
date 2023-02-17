package com.gstsgy.tools.excel.function;

@FunctionalInterface
public interface ReadLine<T> {
    void read(T obj);
}

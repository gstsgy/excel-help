package com.github.excel.help.function;

@FunctionalInterface
public interface ReadLine<T> {
    void read(T obj);
}

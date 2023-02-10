package com.github.excel.help.function;
@FunctionalInterface
public interface ReadLineIncludeOrder<T> {
    void read(int order,T obj);
}

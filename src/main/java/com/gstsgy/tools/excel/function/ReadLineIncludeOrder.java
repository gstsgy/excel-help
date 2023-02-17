package com.gstsgy.tools.excel.function;
@FunctionalInterface
public interface ReadLineIncludeOrder<T> {
    void read(int order,T obj);
}

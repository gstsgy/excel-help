package com.gstsgy.tools.excel;

import java.util.Objects;

public interface APP_CONST {
    String DIR = Objects.requireNonNull(APP_CONST.class.getClassLoader().getResource("")).getPath();

    default String getDir(String fileName){
        return DIR;
    }
}

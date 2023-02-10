package com.github.excel.help.test;

import com.github.excel.help.APP_CONST;
import com.github.excel.help.ReadExcel;
import com.github.excel.help.bean.InOrder;
import com.github.excel.help.bean.SetIgnore;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.lang.reflect.InvocationTargetException;
import java.util.List;

public class ReadDemo {
    public static void main(String[] args) throws IOException, InvocationTargetException, InstantiationException, IllegalAccessException {

        /**  EXCEL 测试数据
         * 序号	编号	数据库表	sql语句	参数
         * 1	1	a	a	a
         * 2	2	b	b	b
         * 3	3	c	c	c
         */
        File file = new File(APP_CONST.DIR+"a.xlsx");

        ReadExcel.ReadExcelBuild<InOrder> readExcelBuild = ReadExcel.build();
        ReadExcel<InOrder> readExcel = readExcelBuild.
                setInputStream(new FileInputStream(file)).
                setTargetType(InOrder.class).
                addSetFuns(new SetIgnore<>()).
                addSetFuns(InOrder::setId).
                addSetFuns(InOrder::setDate).
                addSetFuns(InOrder::setName).
                addSetFuns(InOrder::setDesc).
                crete();
        List<InOrder> data = readExcel.readAll();
        readExcel.readAll();
        System.out.println(data);

    }
}

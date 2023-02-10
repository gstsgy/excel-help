package com.github.excel.help.test;


import com.github.excel.help.APP_CONST;
import com.github.excel.help.WriteExcel;
import com.github.excel.help.bean.DateTimeNow;
import com.github.excel.help.bean.InOrder;
import com.github.excel.help.bean.OrderFunction;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.lang.reflect.InvocationTargetException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

public class WriteDemo2 {
    public static void main(String[] args) throws IOException, InvocationTargetException, IllegalAccessException {
        // 准备数据
        List<InOrder> data = new ArrayList<>();
        for (int i = 0; i < 100; i++) {
            InOrder logDO = new InOrder();
            logDO.setId(i);
            logDO.setDate(getStr());
            logDO.setDesc(getStr());
            logDO.setName(getStr());
            data.add(logDO);
        }
        File file = new File(APP_CONST.DIR+"b.xlsx");

        WriteExcel.WriteExcelBuild<InOrder> writeExcelBuild = WriteExcel.build();


        WriteExcel<InOrder>  writeExcel= writeExcelBuild.
                setOutputStream(new FileOutputStream(file)).
                setTitles(Arrays.asList("序号","编号","数据库表","sql语句","参数","时间")).
                addGetMethos(new OrderFunction<>()).
                addGetMethos(InOrder::getId).
                addGetMethos(InOrder::getDate).
                addGetMethos(InOrder::getName).
                addGetMethos(InOrder::getDesc).
                addGetMethos(new DateTimeNow<>()).
                        create();

        writeExcel.writeToNewSheet(data);
        writeExcel.save();
    }

    public static String getStr(){
        int minLenth = 4; // 定义随机数的最小值
        int maxLenth = 10; // 定义随机数的最大值
        // 产生一个2~100的数
        int strLenth = minLenth + (int) (Math.random() * (maxLenth - minLenth));
        StringBuilder stringBuffer = new StringBuilder();
        for (int i = 0; i <strLenth ; i++) {
            stringBuffer.append((char)(97 +  (int)(Math.random() * (122 - 97))));
        }

        return stringBuffer.toString();

    }
}

package com.gstsgy.tools.excel.test;


import com.gstsgy.tools.excel.APP_CONST;
import com.gstsgy.tools.excel.WriteExcel;
import com.gstsgy.tools.excel.bean.InOrder;
import com.gstsgy.tools.excel.bean.OrderFunction;
import org.junit.jupiter.api.Test;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.lang.reflect.InvocationTargetException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

public class WriteDemo {
    @Test
    public  void t1() throws IOException, InvocationTargetException, IllegalAccessException {
        // 准备数据
        List<InOrder> data = new ArrayList<>();
        for (int i = 0; i < 100; i++) {
            InOrder logDO = new InOrder();
            logDO.setId(i);
            logDO.setName(getStr());
            logDO.setDate(getStr());
            logDO.setDesc(getStr());
            data.add(logDO);
        }
        File file = new File(APP_CONST.DIR+"a.xlsx");
        System.out.println(file.getAbsolutePath());
        WriteExcel.WriteExcelBuild<InOrder> writeExcelBuild = WriteExcel.build();


        WriteExcel<InOrder>  writeExcel= writeExcelBuild.
                setOutputStream(new FileOutputStream(file)).
                //setTargetType(DbLogDO.class).
                setTitles(Arrays.asList("序号","编号","数据库表","sql语句","参数")).
                        addGetMethos(new OrderFunction<>(1000,1)).
                addGetMethos(InOrder::getId).
                addGetMethos(InOrder::getName).
                addGetMethos(InOrder::getDate).
                addGetMethos(InOrder::getDesc).
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

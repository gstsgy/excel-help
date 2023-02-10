package com.github.excel.help;


import com.github.excel.help.function.GetFunction;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;

public class WriteExcel<T> {

    private OutputStream outputStream;

    private String header;

    private List<String> titles;

    private List<GetFunction<T, Object>> getMethos;

    private Workbook workbook;

    private int sheetOrder;

    private Sheet sheet;

    private int currRow;
    private  Map<String, CellStyle> styles ;

    private WriteExcel() {
        this.workbook = new XSSFWorkbook();
    }

    public static <T> WriteExcelBuild<T> build() {
        return new WriteExcelBuild<>(new WriteExcel<>());
    }

    public static class WriteExcelBuild<T> {
        private final WriteExcel<T> writeExcel;

        private WriteExcelBuild(WriteExcel<T> writeExcel) {
            this.writeExcel = writeExcel;
        }

        public WriteExcelBuild<T> setOutputStream(OutputStream outputStream) {
            this.writeExcel.outputStream = outputStream;
            return this;
        }

        public WriteExcelBuild<T> setHeader(String header) {
            this.writeExcel.header = header;
            return this;
        }

        public WriteExcelBuild<T> setTitles(List<String> titles) {
            this.writeExcel.titles = titles;
            return this;
        }

        public WriteExcelBuild<T> setGetMethos(List<GetFunction<T, Object>> getMethos) {
            this.writeExcel.getMethos = getMethos;
            return this;
        }

        public WriteExcelBuild<T> addGetMethos(GetFunction<T, Object> fn) {

            if (this.writeExcel.getMethos == null) {
                this.writeExcel.getMethos = new ArrayList<>();
            }
            this.writeExcel.getMethos.add(fn);
            return this;
        }


        public WriteExcel<T> create() throws IOException {
            if (this.writeExcel.outputStream == null) {
                throw new IOException("outputStream 必须设置");
            }

            if (this.writeExcel.getMethos == null) {
                throw new RuntimeException("未指定导出字段");
            }
            if (this.writeExcel.titles == null) {
                this.writeExcel.titles = this.writeExcel.getMethos.stream().map(GetFunction::getFieldName).collect(Collectors.toList());
            }
            if (this.writeExcel.titles.size() != this.writeExcel.getMethos.size()) {
                throw new RuntimeException("titles 与getMethos 长度必须相等");
            }
            return this.writeExcel;
        }


    }

    public WriteExcelBuild<T> refresh() {
        return new WriteExcelBuild<>(this);
    }


    public void writeToNewSheet(List<T> data) {
        initSheet("sheet" + sheetOrder++,data.size());
        currRow = 1;
        writeExcelSheet( data);
    }

    public void append(List<T> data) {
        writeExcelSheet(data);
    }

    private void initSheet(String sheetName,int dataSize){
        // 判断数据量大小
        if (dataSize > 1000) {
            workbook = new SXSSFWorkbook();
        }

        styles = createStyles(workbook);
        // 生成一个表格
        sheet = workbook.createSheet(sheetName);

        // 设置表格默认列宽度为15个字节
        sheet.setDefaultColumnWidth((short) 15);

        // 创建标题行
        Row row = sheet.createRow(0);
        for (int i = 0; i < titles.size(); i++) {
            Cell cell = row.createCell(i);
            cell.setCellStyle(styles.get("title"));
            cell.setCellValue(titles.get(i));
        }
    }


    /**
     * 创建Excel文件
     *
     */
    private void writeExcelSheet(List<T> data) {

        // 写入正文
        for (T datum : data) {
           Row row = sheet.createRow(currRow);
            if (datum == null) {
                continue;
            }
            for (int i1 = 0; i1 < getMethos.size(); i1++) {

                // 在指定序号处创建cell
                Cell cell = row.createCell(i1);
                // 设置cell的样式
                if (currRow % 2 == 1) {
                    cell.setCellStyle(styles.get("cellA"));
                } else {
                    cell.setCellStyle(styles.get("cellB"));
                }
                Object object = getMethos.get(i1).getVal(datum);
                if (object != null) {
                    cell.setCellValue(object.toString());
                }
            }
            currRow++;
        }
    }

    public void save() throws IOException {
        try {
            //outputStream = new FileOutputStream(filepath);
            workbook.write(outputStream);

        } finally {
            if (outputStream != null) {
                outputStream.close();
            }
            workbook.close();
        }
    }

    /**
     * 设置格式
     */
    private static Map<String, CellStyle> createStyles(Workbook wb) {
        Map<String, CellStyle> styles = new HashMap<>();

        // 标题样式
        XSSFCellStyle titleStyle = (XSSFCellStyle) wb.createCellStyle();
        titleStyle.setAlignment(HorizontalAlignment.CENTER); // 水平对齐
        titleStyle.setVerticalAlignment(VerticalAlignment.CENTER); // 垂直对齐
        titleStyle.setLocked(true); // 样式锁定
        titleStyle.setFillForegroundColor(IndexedColors.LIGHT_YELLOW.getIndex());
        Font titleFont = wb.createFont();
        titleFont.setFontHeightInPoints((short) 16);
        titleFont.setBold(true);
        titleFont.setFontName("微软雅黑");
        titleStyle.setFont(titleFont);
        styles.put("title", titleStyle);

        // 文件头样式
        XSSFCellStyle headerStyle = (XSSFCellStyle) wb.createCellStyle();
        headerStyle.setAlignment(HorizontalAlignment.CENTER);
        headerStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        headerStyle.setFillForegroundColor(IndexedColors.LIGHT_BLUE.getIndex()); // 前景色
        headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND); // 颜色填充方式
        headerStyle.setWrapText(true);
        headerStyle.setBorderRight(BorderStyle.THIN); // 设置边界
        headerStyle.setRightBorderColor(IndexedColors.BLACK.getIndex());
        headerStyle.setBorderLeft(BorderStyle.THIN);
        headerStyle.setLeftBorderColor(IndexedColors.BLACK.getIndex());
        headerStyle.setBorderTop(BorderStyle.THIN);
        headerStyle.setTopBorderColor(IndexedColors.BLACK.getIndex());
        headerStyle.setBorderBottom(BorderStyle.THIN);
        headerStyle.setBottomBorderColor(IndexedColors.BLACK.getIndex());
        Font headerFont = wb.createFont();
        headerFont.setFontHeightInPoints((short) 12);
        headerFont.setColor(IndexedColors.WHITE.getIndex());
        titleFont.setFontName("微软雅黑");
        headerStyle.setFont(headerFont);
        styles.put("header", headerStyle);

        Font cellStyleFont = wb.createFont();
        cellStyleFont.setFontHeightInPoints((short) 12);
        cellStyleFont.setColor(IndexedColors.BLUE_GREY.getIndex());
        cellStyleFont.setFontName("微软雅黑");

        // 正文样式A
        XSSFCellStyle cellStyleA = (XSSFCellStyle) wb.createCellStyle();
        cellStyleA.setAlignment(HorizontalAlignment.CENTER); // 居中设置
        cellStyleA.setVerticalAlignment(VerticalAlignment.CENTER);
        cellStyleA.setWrapText(true);
        cellStyleA.setBorderRight(BorderStyle.THIN);
        cellStyleA.setRightBorderColor(IndexedColors.BLACK.getIndex());
        cellStyleA.setBorderLeft(BorderStyle.THIN);
        cellStyleA.setLeftBorderColor(IndexedColors.BLACK.getIndex());
        cellStyleA.setBorderTop(BorderStyle.THIN);
        cellStyleA.setTopBorderColor(IndexedColors.BLACK.getIndex());
        cellStyleA.setBorderBottom(BorderStyle.THIN);
        cellStyleA.setBottomBorderColor(IndexedColors.BLACK.getIndex());
        cellStyleA.setFont(cellStyleFont);
        styles.put("cellA", cellStyleA);

        // 正文样式B:添加前景色为浅黄色
        XSSFCellStyle cellStyleB = (XSSFCellStyle) wb.createCellStyle();
        cellStyleB.setAlignment(HorizontalAlignment.CENTER);
        cellStyleB.setVerticalAlignment(VerticalAlignment.CENTER);
        cellStyleB.setFillForegroundColor(IndexedColors.LIGHT_YELLOW.getIndex());
        cellStyleB.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        cellStyleB.setWrapText(true);
        cellStyleB.setBorderRight(BorderStyle.THIN);
        cellStyleB.setRightBorderColor(IndexedColors.BLACK.getIndex());
        cellStyleB.setBorderLeft(BorderStyle.THIN);
        cellStyleB.setLeftBorderColor(IndexedColors.BLACK.getIndex());
        cellStyleB.setBorderTop(BorderStyle.THIN);
        cellStyleB.setTopBorderColor(IndexedColors.BLACK.getIndex());
        cellStyleB.setBorderBottom(BorderStyle.THIN);
        cellStyleB.setBottomBorderColor(IndexedColors.BLACK.getIndex());
        cellStyleB.setFont(cellStyleFont);
        styles.put("cellB", cellStyleB);

        return styles;
    }
}

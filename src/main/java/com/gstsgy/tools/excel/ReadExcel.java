
package com.gstsgy.tools.excel;

import com.gstsgy.tools.excel.bean.SetIgnore;
import com.gstsgy.tools.excel.function.ReadLine;
import com.gstsgy.tools.excel.function.SetFunction;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;

import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.InvocationTargetException;
import java.math.BigDecimal;
import java.util.*;

import static org.apache.poi.ss.usermodel.Cell.CELL_TYPE_STRING;

public class ReadExcel<T> {

    private Class<T> targetType;

    private List<SetFunction<T,?>> setFuns;

    private InputStream inputStream;

    private Workbook workbook;

    private ReadExcel() {
        wrap2Base.put(Integer.class, int.class);
        wrap2Base.put(Byte.class, byte.class);
        wrap2Base.put(Short.class, short.class);
        wrap2Base.put(Long.class, long.class);

        wrap2Base.put(Float.class, float.class);
        wrap2Base.put(Double.class, double.class);

        wrap2Base.put(Boolean.class, boolean.class);
        wrap2Base.put(Character.class, char.class);
    }

    public static Workbook read(InputStream inputStream) throws IOException {
        Workbook workbook = null;
        try {

            workbook = WorkbookFactory.create(inputStream);
        } catch (IOException | InvalidFormatException e) {
            e.printStackTrace();
        } finally {
            if (inputStream != null) {
                inputStream.close();
            }
            if (workbook != null) {
                workbook.close();
            }
        }
        return workbook;
    }

    public static <T> ReadExcelBuild<T> build() {
        return new ReadExcelBuild<>(new ReadExcel<>());
    }


    public ReadExcelBuild<T> refresh() {
        return new ReadExcelBuild<T>(this);
    }

    public List<T> readAll() throws EncryptedDocumentException, InvocationTargetException, InstantiationException, IllegalAccessException {
        List<T> list = new ArrayList<>();
        Iterator<Sheet> sheetIterator = workbook.sheetIterator();
        while (sheetIterator.hasNext()) {
            list.addAll(readSheel(sheetIterator.next()));
        }
        return list;
    }

    public List<T> read(int sheetIndex)
            throws EncryptedDocumentException, InvocationTargetException, InstantiationException, IllegalAccessException {
        Sheet sheet = workbook.getSheetAt(sheetIndex);
        return readSheel(sheet);
    }


    private List<T> readSheel(Sheet sheet)
            throws EncryptedDocumentException, InstantiationException, IllegalAccessException {
        if (sheet != null) {
            String[] titles = null;
            int rowNos = sheet.getLastRowNum();// ??????excel??????????????????
            List<T> dataArr = new ArrayList<>(rowNos);
            for (int i = 0; i <= rowNos; i++) {// ?????????
                Row row = sheet.getRow(i);
                if (row != null) {
                    if (titles == null) {
                        titles = readTitle(row);
                    } else {
                        dataArr.add(readLine(row));
                    }
                }
            }
            return dataArr;
        }
        return new ArrayList<>();
    }


    private String[] readTitle(Row row) {
        int columNos = row.getLastCellNum();// ?????????????????????

        String[] titles = new String[columNos];
        for (int j = 0; j < columNos; j++) {
            Cell cell = row.getCell(j);


            if (cell != null) {
                cell.setCellType(CELL_TYPE_STRING);
                titles[j] = cell.getStringCellValue();
            }
        }
        return titles;
    }

    /**
     * ????????????
     */

    private T readLine(Row row) throws InstantiationException, IllegalAccessException {
        int columNos = row.getLastCellNum();// ?????????????????????
        T obj = targetType.newInstance();


        for (int j = 0; j < columNos; j++) {
            Cell cell = row.getCell(j);
            cell.setCellType(CELL_TYPE_STRING);
            SetFunction<T,?> setFunction = this.setFuns.get(j);
            if(setFunction instanceof SetIgnore){
                continue;
            }
            Object val = cell.getStringCellValue();
            ((SetFunction<T,Object>) setFunction).setVal(obj, convert(val,setFunction.getImplMethodParamType().get(1)));
        }
        return obj;
    }

    private void  readLine(Row row , ReadLine<T> readLine) throws InstantiationException, IllegalAccessException {
        int columNos = row.getLastCellNum();// ?????????????????????
        T obj = targetType.newInstance();


        for (int j = 0; j < columNos; j++) {
            Cell cell = row.getCell(j);
            cell.setCellType(CELL_TYPE_STRING);
            SetFunction<T,?> setFunction = this.setFuns.get(j);
            if(setFunction instanceof SetIgnore){
                continue;
            }
            Object val = cell.getStringCellValue();
            ((SetFunction<T,Object>) setFunction).setVal(obj, convert(val,setFunction.getImplMethodParamType().get(1)));
        }
        readLine.read(obj);
    }


    public static class ReadExcelBuild<T> {
        private final ReadExcel<T> readExcel;

        private ReadExcelBuild(ReadExcel<T> readExcel) {
            this.readExcel = readExcel;
        }

        public ReadExcelBuild<T> setTargetType(Class<T> targetType) {
            this.readExcel.targetType = targetType;
            return this;
        }

        public ReadExcelBuild<T> setSetFuns(List<SetFunction<T, ?>> setFuns) {
            this.readExcel.setFuns = setFuns;
            return this;
        }

        public ReadExcelBuild<T> setInputStream(InputStream inputStream) {
            this.readExcel.inputStream = inputStream;
            return this;
        }

        public<P> ReadExcelBuild<T> addSetFuns(SetFunction<T, P> setFunction) {
            if (this.readExcel.setFuns == null) {
                this.readExcel.setFuns = new ArrayList<>();
            }
            this.readExcel.setFuns.add(setFunction);
            return this;
        }


        public ReadExcel<T> crete() throws IOException {
            if (this.readExcel.inputStream == null) {
                throw new IOException("inputStream ????????????");
            }

            if (this.readExcel.targetType == null) {
                throw new RuntimeException("?????????????????????");
            }
            if (this.readExcel.setFuns == null) {
                throw new RuntimeException("set???????????????");
            }

            try {

                this.readExcel.workbook = WorkbookFactory.create(this.readExcel.inputStream);
            } catch (IOException | InvalidFormatException e) {
                e.printStackTrace();
            } finally {
                if (this.readExcel.inputStream != null) {
                    this.readExcel.inputStream.close();
                }
                if (this.readExcel.workbook != null) {
                    this.readExcel.workbook.close();
                }
            }
            return this.readExcel;
        }
    }
    public List<Class<?>> baseClass = Arrays.asList(int.class, byte.class, short.class, long.class, float.class, double.class, boolean.class, char.class);
    public List<Class<?>> baseWrapClass = Arrays.asList(Integer.class, Byte.class, Short.class, Long.class,
            Float.class, Double.class, Boolean.class, Character.class);
    public Map<Class, Class> wrap2Base = new HashMap<>();




    /**
     * ????????????
     *
     * @param value       ????????????
     * @param targetClass ????????????
     * @return ???????????????
     */
    public  Object convert(Object value, Class targetClass) {
        if (value == null) {
            return null;
        }
        if (targetClass == String.class && value.getClass() == String.class) {
            return value;
        }
        if (targetClass == String.class) {
            return String.valueOf(value);
        }
        // ???????????????
        if (targetClass == value.getClass() && this.baseWrapClass.contains(targetClass)) {
            return value;
        }
        // ??????????????????
        if (targetClass == value.getClass() && this.baseClass.contains(targetClass)) {
            return value;
        }
        // target ???????????? ????????? value.getClass() ????????????
        if (targetClass == this.wrap2Base.get(value.getClass()) && this.baseWrapClass.contains(value.getClass())) {
            return value;
        }
        // value.getClass() ???????????? ????????? target ????????????
        if (value.getClass() == this.wrap2Base.get(targetClass) && this.baseWrapClass.contains(targetClass)) {
            return value;
        }
//        // ????????????
//        if (Arrays.asList(targetClass.getInterfaces()).contains(BaseEnum.class)) {
//            UniversalEnumConverterFactory converterFactory = new UniversalEnumConverterFactory();
//            org.springframework.core.convert.converter.Converter<String, BaseEnum> baseEnumConverter = converterFactory.getConverter1(targetClass);
//            return baseEnumConverter.convert(String.valueOf(value));
//        }
        // value ???string target????????????
        if (targetClass == Integer.class || targetClass == int.class) {
            return Integer.valueOf(String.valueOf(value));

        } else if (targetClass == Long.class || targetClass == long.class) {
            return Long.valueOf(String.valueOf(value));
        } else if (targetClass == Double.class || targetClass == double.class) {
            return Double.valueOf(String.valueOf(value));
        } else if (targetClass == Float.class || targetClass == float.class) {
            return Float.valueOf(String.valueOf(value));
        } else if (targetClass == Byte.class || targetClass == byte.class) {
            return Byte.valueOf(String.valueOf(value));
        } else if (targetClass == Short.class || targetClass == short.class) {
            return Short.valueOf(String.valueOf(value));
        } else if (targetClass == Character.class || targetClass == char.class) {
            return String.valueOf(value).charAt(0);
        } else if (targetClass == Boolean.class || targetClass == boolean.class) {
            return Boolean.valueOf(String.valueOf(value));
        } else if (targetClass == BigDecimal.class) {
            return new BigDecimal(String.valueOf(value));
        }
        throw new RuntimeException(String.format("?????????????????????value?????????%s,???????????????%s",value.getClass().getName(),targetClass.getName()));
    }
}

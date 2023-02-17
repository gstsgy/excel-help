package com.gstsgy.tools.excel.function;

/**
 *
 * @param <T> 目标对象类型
 * @param <R> 目标对象get返回类型
 */

@FunctionalInterface
public interface GetFunction<T,R> extends BaseFunction{

    R getVal(T t);

    default String getFieldName(){
        String implMethodName = getMethodName();
        // 确保方法是符合规范的get方法，boolean类型是is开头
        if (!implMethodName.startsWith("is") && !implMethodName.startsWith("get")) {
            throw new RuntimeException("get方法名称: " + implMethodName + ", 不符合java bean规范");
        }

        // get方法开头为 is 或者 get，将方法名 去除is或者get，然后首字母小写，就是属性名
        int prefixLen = implMethodName.startsWith("is") ? 2 : 3;

        String fieldName = implMethodName.substring(prefixLen);
        String firstChar = fieldName.substring(0, 1);
        fieldName = fieldName.replaceFirst(firstChar, firstChar.toLowerCase());


        return fieldName;
    }

    default int getParamsCount(){
        return 0;
    }
}

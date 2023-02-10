package com.github.excel.help.function;

import java.io.Serializable;
import java.lang.invoke.SerializedLambda;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.util.Arrays;
import java.util.List;
import java.util.Objects;
import java.util.stream.Collectors;

/**
 * @author guyue
 * @version 3.0
 * @description: TODO
 * @date 2021/11/8 上午9:53
 */
public interface BaseFunction extends Serializable {

    int getParamsCount();

    default String getMethodName() {
        return getSerializedLambda().getImplMethodName();
    }

    default String getClassName() {
        return getSerializedLambda().getImplClass().replace("/", ".");
    }

    default SerializedLambda getSerializedLambda() {
        Method writeReplaceMethod;
        //Method[] c = fn.getClass().getDeclaredMethods();

        try {
            writeReplaceMethod = this.getClass().getDeclaredMethod("writeReplace");
        } catch (NoSuchMethodException e) {
            throw new RuntimeException(e);
        }
        // 从序列化方法取出序列化的lambda信息
        boolean isAccessible = writeReplaceMethod.isAccessible();
        writeReplaceMethod.setAccessible(true);
        SerializedLambda serializedLambda;
        try {
            serializedLambda = (SerializedLambda) writeReplaceMethod.invoke(this);
        } catch (IllegalAccessException | InvocationTargetException e) {
            throw new RuntimeException(e);
        }

        return serializedLambda;
    }


    default List<Class> getImplMethodParamType() {
        SerializedLambda lambda = getSerializedLambda();
        String methodType = lambda.getInstantiatedMethodType();
        methodType = methodType.substring(methodType.indexOf("(") + 1, methodType.indexOf(")"));
        String[] className = methodType.split(";");
        return Arrays.stream(className).filter(Objects::nonNull).map(it -> it.substring(1).replace("/", ".")).
                map(it -> {
                    try {
                        return Class.forName(it);
                    } catch (ClassNotFoundException e) {
                        e.printStackTrace();
                        return null;
                    }
                }).collect(Collectors.toList());
    }
}

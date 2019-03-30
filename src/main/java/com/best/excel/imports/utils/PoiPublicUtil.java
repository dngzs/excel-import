package com.best.excel.imports.utils;

import com.best.excel.annotation.ExcelCollection;
import com.best.excel.annotation.ExcelTitle;
import org.apache.commons.lang3.StringUtils;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.lang.reflect.Field;
import java.math.BigDecimal;
import java.util.*;

public class PoiPublicUtil {

    private static final Logger LOGGER = LoggerFactory.getLogger(PoiPublicUtil.class);

    private PoiPublicUtil() {

    }


    /**
     * 彻底创建一个对象
     *
     * @param clazz
     * @return
     */
    public static Object createObject(Class<?> clazz, String targetId) {
        Object obj = null;
        try {
            if (clazz.equals(Map.class)) {
                return new LinkedHashMap<String,Object>();
            }
            obj = clazz.newInstance();
            Field[] fields = getClassFields(clazz);
            for (Field field : fields) {
                if (isNotUserExcelUserThis(null, field, targetId)) {
                    continue;
                }
                if (isCollection(field.getType())) {
                    ExcelCollection collection = field.getAnnotation(ExcelCollection.class);
                    PoiReflectorUtil.fromCache(clazz).setValue(obj, field.getName(),
                            collection.type().newInstance());
                } else if (!isJavaClass(field) && !field.getType().isEnum()) {
                    PoiReflectorUtil.fromCache(clazz).setValue(obj, field.getName(),
                            createObject(field.getType(), targetId));
                }
            }

        } catch (Exception e) {
            LOGGER.error(e.getMessage(), e);
            throw new RuntimeException("创建对象异常");
        }
        return obj;

    }

    /**
     * 判断是不是集合的实现类
     *
     * @param clazz
     * @return
     */
    public static boolean isCollection(Class<?> clazz) {
        return Collection.class.isAssignableFrom(clazz);
    }



    /**
     * 获取class的字段（包括父类的）
     *
     * @param clazz
     * @return
     */
    public static Field[] getClassFields(Class<?> clazz) {
        List<Field> list = new ArrayList<Field>();
        Field[] fields;
        do {
            fields = clazz.getDeclaredFields();
            for (int i = 0; i < fields.length; i++) {
                list.add(fields[i]);
            }
            clazz = clazz.getSuperclass();
        } while (clazz != Object.class && clazz != null);
        return list.toArray(fields);
    }


    /**
     * 判断是否不要在这个excel操作中
     *
     * @param exclusionsList
     * @param field
     * @param targetId
     * @return
     */
    public static boolean isNotUserExcelUserThis(List<String> exclusionsList, Field field,
                                                 String targetId) {
        boolean boo = true;
         if (boo && field.getAnnotation(ExcelTitle.class) != null
                && isUseInThis(field.getAnnotation(ExcelTitle.class).name(), targetId)
                && (exclusionsList == null
                || !exclusionsList.contains(field.getAnnotation(ExcelTitle.class).name()))) {
            boo = false;
        }
        return boo;
    }

    /**
     * 判断是不是使用
     *
     * @param exportName
     * @param targetId
     * @return
     */
    private static boolean isUseInThis(String exportName, String targetId) {
        return targetId == null || "".equals(exportName) || exportName.indexOf("_") < 0
                || exportName.indexOf(targetId) != -1;
    }


    /**
     * 统一 key的获取规则
     * @param key
     * @param targetId
     * @return
     */
    public static String getValueByTargetId(String key, String targetId, String defalut) {
        if (StringUtils.isEmpty(targetId) || key.indexOf("_") < 0) {
            return key;
        }
        String[] arr = key.split(",");
        String[] tempArr;
        for (String str : arr) {
            tempArr = str.split("_");
            if (tempArr == null || tempArr.length < 2) {
                return defalut;
            }
            if (targetId.equals(tempArr[1])) {
                return tempArr[0];
            }
        }
        return defalut;
    }

    /**
     * 是不是java基础类
     *
     * @param field
     * @return
     */
    public static boolean isJavaClass(Field field) {
        Class<?> fieldType = field.getType();
        boolean isBaseClass = false;
        if (fieldType.isArray()) {
            isBaseClass = false;
        } else if (fieldType.isPrimitive() || fieldType.getPackage() == null
                || "java.lang".equals(fieldType.getPackage().getName())
                || "java.math".equals(fieldType.getPackage().getName())
                || "java.sql".equals(fieldType.getPackage().getName())
                || "java.util".equals(fieldType.getPackage().getName())) {
            isBaseClass = true;
        }
        return isBaseClass;
    }


    /**
     * double to String 防止科学计数法
     * @param value
     * @return
     */
    public static String doubleToString(Double value) {
        String temp = value.toString();
        if (temp.contains("E")) {
            BigDecimal bigDecimal = new BigDecimal(temp);
            temp = bigDecimal.toPlainString();
        }
        return temp;
    }
}

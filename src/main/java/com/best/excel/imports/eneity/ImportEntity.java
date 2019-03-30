package com.best.excel.imports.eneity;


import java.lang.reflect.Method;
import java.util.List;

/**
 * excel 导入工具类,对cell类型做映射
 *
 * @author JueYue
 * @version 1.0 2013年8月24日
 */

public class ImportEntity {

    /**
     * 时间格式化
     */
    private String format;

    /**
     * 文本标识
     */
    private int type;

    /**
     * 表头
     */
    private String name;

    /**
     * 枚举导入静态方法
     */
    private String enumImportMethod;

    /**
     * get set方法
     */
    private List<Method> methods;

    /**
     * set/get方法
     */
    private Method method;

    public List<Method> getMethods() {
        return methods;
    }

    public Method getMethod() {
        return method;
    }

    public void setMethod(Method method) {
        this.method = method;
    }

    public void setMethods(List<Method> methods) {
        this.methods = methods;
    }

    public String getEnumImportMethod() {
        return enumImportMethod;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public int getType() {
        return type;
    }

    public void setType(int type) {
        this.type = type;
    }

    public String getFormat() {
        return format;
    }

    public void setFormat(String format) {
        this.format = format;
    }

    public void setEnumImportMethod(String enumImportMethod) {
        this.enumImportMethod = enumImportMethod;
    }
}


package com.best.excel.imports.eneity;


import java.lang.reflect.Method;
import java.util.List;

/**
 * excel 导入工具类,对cell类型做映射
 * @author JueYue
 * @version 1.0 2013年8月24日
 */

public class ImportEntity {

    private String format;

    private int type;

    private String name;

    /**
     * 对应 Collection NAME
     */
    private String                  collectionName;

    /**
     * 对应exportType
     */
    private String                  classType;
    /**
     * 后缀
     */
    private String                  suffix;
    /**
     * 导入校验字段
     */
    private boolean                 importField;

    /**
     * 枚举导入静态方法
     */
    private String                   enumImportMethod;

    private List<ImportEntity> list;


    private List<Method> methods;

    /**
     * set/get方法
     */
    private Method   method;

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

    public String getClassType() {
        return classType;
    }

    public String getCollectionName() {
        return collectionName;
    }

    public List<ImportEntity> getList() {
        return list;
    }

    public void setClassType(String classType) {
        this.classType = classType;
    }

    public void setCollectionName(String collectionName) {
        this.collectionName = collectionName;
    }

    public void setList(List<ImportEntity> list) {
        this.list = list;
    }


    public String getSuffix() {
        return suffix;
    }

    public void setSuffix(String suffix) {
        this.suffix = suffix;
    }

    public boolean isImportField() {
        return importField;
    }

    public void setImportField(boolean importField) {
        this.importField = importField;
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


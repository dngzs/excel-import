package com.best.excel.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * excle表头标题
 */
@Retention(RetentionPolicy.RUNTIME)
@Target({ElementType.FIELD})
public @interface ExcelTitle {

    String name();

    /**
     * 导入类型 1 是文本
     */
    int type() default 1;

    /**
     * 如果是时间，必须要设置
     *
     * @return
     */
    String format() default "";

    /**
     * 枚举导入使用的函数
     *
     * @return
     */
    String enumImportMethod() default "";

}

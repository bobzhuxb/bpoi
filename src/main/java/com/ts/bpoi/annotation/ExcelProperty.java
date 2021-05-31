package com.ts.bpoi.annotation;

import java.lang.annotation.*;

/**
 * Excel解析
 * @author Bob
 */
@Documented
@Target({ElementType.FIELD})
@Retention(RetentionPolicy.RUNTIME)
public @interface ExcelProperty {

    // 列标题
    String value();

    // 排序
    int index();

    // 验证正则
    String regex() default ".*";

    // 是否允许空值
    boolean nullable() default true;

    // 空值时设定的默认值（该值为非空字符串时有效，且忽略nullable的影响）
    String emptyToDefault() default "";

    // 日期格式
    String dateFormat() default "yyyy-MM-dd";

    // 取值范围
    String[] optionList() default {};

}

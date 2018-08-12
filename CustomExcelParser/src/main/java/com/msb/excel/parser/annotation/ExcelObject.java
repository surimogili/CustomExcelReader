package com.msb.excel.parser.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

@Retention(value = RetentionPolicy.RUNTIME)
@Target({ElementType.TYPE})
public @interface ExcelObject {

    
    ParseType parseType();

    int start();

    int end() default 0;

    boolean loop() default false;

    boolean ignoreAllZerosOrNullRows() default false;
    
    int looplength() default 0;
}

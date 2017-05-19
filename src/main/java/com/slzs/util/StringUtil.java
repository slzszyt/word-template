package com.slzs.util;

/**
 * String工具类
 * 记录一些常用的针对字符串的通用处理函数
 * @author slzs
 * 
 */
public class StringUtil {

    /**
     * 验证字符串类型的参数是否为空
     * 
     * @param strs
     * @return boolean  多个同时为空 true 
     */
    public static boolean isEmpty(String... strs) {
        boolean b = false;
        if (strs == null) {
            b = true;
        } else {
            b = true;
            for (String str : strs) {
                b = b && (str == null || "".equals(str));
            }
        }
        return b;
    }

    /**
     * 验证字符串类型的参数是否不为空
     * 
     * @param strs
     * @return boolean  多个同时不为空 true 
     */
    public static boolean isNotEmpty(String... strs) {
        boolean b = false;
        if (strs == null) {
            b = false;
        } else {
            b = true;
            for (String str : strs) {
                b = b && !isEmpty(str);
            }
        }
        return b;
    }

    /**
     * 验证字符串类型的参数是否为空(去除首尾空格)
     * 
     * @param strs 多个同时trim后为空
     * @return boolean
     */
    public static boolean isEmptyTrim(String... strs) {
        return isEmpty(trim(strs));
    }

    /**
     * 验证字符串类型的参数是否不为空(去除首尾空格)
     * 
     * @param strs  多个同时trim后不为空
     * @return boolean
     */
    public static boolean isNotEmptyTrim(String... strs) {
        return isNotEmpty(trim(strs));
    }

    /**
     * 字符串数组值trim
     * @author: slzs
     * 2013-12-23 下午4:26:31
     * @param strs
     * @return String[] 返回trim后的字符串数组
     * 
     */
    public static String[] trim(String... strs) {
        if (strs != null && strs.length > 0) {
            for (int i = 0; i < strs.length; i++) {
                if (strs[i] != null) {
                    strs[i] = strs[i].trim();
                }
            }
        }
        return strs;
    }

}

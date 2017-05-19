package com.slzs.util;

import java.lang.reflect.Array;
import java.util.Collection;
import java.util.Map;

public class ObjectUtil {

    /**
     * 检查数据为空（为null或者初始值：Map size 0,list size=0,[] length=0,number 0,String trim ""）
     * @author: slzs
     * 2013-12-25 上午10:29:27
     * @param checkObj Object/List/Map/number/String/...
     * @return boolean 为空 true
     * 
     */
    public static boolean isEmpty(Object checkObj) {
        boolean isEmpty = false;
        if (checkObj != null) {
            if (checkObj instanceof String) { // 字符串
                isEmpty = StringUtil.isEmptyTrim(String.valueOf(checkObj));
            } else if (checkObj instanceof Map) { // map
                isEmpty = ((Map<?, ?>) checkObj).isEmpty();
            } else if (checkObj instanceof Collection) { // 集合
                isEmpty = ((Collection<?>) checkObj).isEmpty();
            } else if (checkObj instanceof Number) {//数值
                isEmpty = ((Number) checkObj).hashCode() == 0;
            } else if (checkObj.getClass().isArray()) {// 数组
                isEmpty = Array.getLength(checkObj) == 0;
            }/* else if(checkObj instanceof JSONNull){
                isEmpty = true;
            }*/
        } else {
            isEmpty = true;
        }
        return isEmpty;
    }

    /**
     * 检查数据非空与empty相反
     * @author: slzs
     * 2013-12-25 上午10:29:27
     * @param checkObj Object/List/Map/number/String/...
     * @return boolean
     * 
     */
    public static boolean isNotEmpty(Object checkObj) {
        return !isEmpty(checkObj);
    }

    /**
     * 计算长度(数组长度、list长度、map长度，其它类型转String，null为0)
     * @author: slzs
     * 2014-3-3 下午5:11:11
     * @param checkObj 所有对象
     * @return int (数组长度、list长度、map长度，其它类型转String，null为0)
     * 
     */
    public static int length(Object checkObj) {
        int length = 0;
        if (checkObj != null) {
            if (checkObj instanceof Map) { // map
                length = ((Map<?, ?>) checkObj).size();
            } else if (checkObj instanceof Collection) { // 集合
                length = ((Collection<?>) checkObj).size();
            } else if (checkObj.getClass().isArray()) {// 数组
                length = Array.getLength(checkObj);
            } else { // 其它转字符串
                length = String.valueOf(checkObj).length();
            }
        }
        return length;
    }
}

package com.xuxueli.poi.excel;

//import org.apache.commons.collections.CollectionUtils;

import org.apache.commons.collections4.CollectionUtils;

import java.util.Collection;
import java.util.Optional;
import java.util.function.Supplier;
import java.util.stream.Stream;

/**
 * @author：shihuiwen
 * @date：Created in 2019/12/4 16:37
 * @description：用于安全处理的工具类，目前主要用于解决空指针异常
 * @modified By：
 * @version:1.0$
 */
public class SafeUtils {
    private SafeUtils() {
    }

    /**
     * 根据一个方法 或者该方法的返回值
     *
     * <p>
     * 主要解决连续get的时候忽略 空指针异常
     *
     * <code>
     * Outter outter = new Outter();
     * //这里返回的值可能为空
     * String name = SafeUtils.safeGet(() -> outter.getInnerL1().getInnerL2().getName());
     * System.out.Println(name);
     * </code>
     *
     * @param resolver 传入的Supplier
     * @return 返回最终的get的值，如果中间对象为空 则返回空
     */
    public static <T> T safeGet(Supplier<T> resolver) {
        try {
            return resolver.get();
        } catch (NullPointerException e) {
            return null;
        }
    }

    /**
     * 根据一个方法 返回Optional
     *
     * <p>
     * 主要解决连续get的时候忽略 空指针异常
     *
     * <code>
     * Outter outter = new Outter();
     * Optional<String> optional = SafeUtils.safeGetOptionally(() -> outter.getInnerL1().getInnerL2().getName());
     * System.out.Println(optional.get());
     * </code>
     *
     * @param resolver 传入的Supplier
     * @return 返回最终的Optional，如果中间对象为空 则返回空的Optional
     */
    public static <T> Optional<T> safeGetOptionally(Supplier<T> resolver) {
        try {
            return Optional.ofNullable(resolver.get());
        } catch (NullPointerException e) {
            return Optional.empty();
        }
    }


    /**
     * 根据集合获取 一个stream 防止出现空指针异常
     *
     * <code>
     * List<String> list = null;
     * SafeUtils.safeGetStream(list).forEach(System.out::println);
     * </code>
     *
     * @param collection 传入一个集合
     * @return 返回一个stream 如果入参为空 则返回一个空的Stream
     */
    public static <T> Stream<T> safeGetStream(Collection<T> collection) {
        if (CollectionUtils.isEmpty(collection)) {
            return Stream.empty();
        }
        return collection.stream();
    }
}

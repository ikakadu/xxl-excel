package com.xuxueli.poi.excel.test.model;

import com.xuxueli.poi.excel.annotation.ExcelField;
import com.xuxueli.poi.excel.annotation.ExcelSheet;
import lombok.Data;
import org.apache.poi.hssf.util.HSSFColor;

import java.util.Date;

/**
 * t_holiday_record
 * 根据Table [t_holiday_record]生成
 */
@Data
@ExcelSheet(name = "节假日列表")
public class HolidayRecord {
    /**
     * Table:     t_holiday_record
     * Column:    id
     * Length:  19
     */
    @ExcelField(name = "序号")
    private Long id;

    /**
     * 区域
     *
     * Table:     t_holiday_record
     * Column:    region
     * Length:  32
     */
    @ExcelField(name = "币种")
    private String currency;

    /**
     * 当地日期
     *
     * Table:     t_holiday_record
     * Column:    local_date
     * Length:  10
     */
    @ExcelField(name = "日期", dateformat = "yyyy/MM/dd")
    private Date localDate;

    /**
     * 节假日时间段，例如：00:00:00-23:59:59
     *
     * Table:     t_holiday_record
     * Column:    time_range
     * Length:  32
     */
    @ExcelField(name = "时间")
    private String timeRange;

    /**
     * 节假日的类型，1：节假日（节假日优先于周末）    ，  2：周末（不包含调休上班日）
     *
     * Table:     t_holiday_record
     * Column:    holiday_type
     * Length:  2
     */
    @ExcelField(name = "类型")
    private String holidayType;

    /**
     * 时区
     *
     * Table:     t_holiday_record
     * Column:    time_zone
     * Length:  12
     */
    @ExcelField(name = "时区")
    private String timeZone;

    /**
     * 编辑人
     *
     * Table:     t_holiday_record
     * Column:    editor
     * Length:  32
     */
    private String editor;

    /**
     * 创建时间
     *
     * Table:     t_holiday_record
     * Column:    gmt_create
     * Length:  19
     */
    private Date gmtCreate;

    /**
     * 更新时间
     *
     * Table:     t_holiday_record
     * Column:    gmt_update
     * Length:  19
     */
    private Date gmtUpdate;
}
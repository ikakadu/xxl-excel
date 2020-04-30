package com.xuxueli.poi.excel;


import com.xuxueli.poi.excel.dto.HolidayRecord;
import org.apache.commons.lang3.StringUtils;
import org.apache.commons.lang3.time.DateUtils;
import org.apache.poi.ss.usermodel.Workbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;



import java.io.*;
import java.net.HttpURLConnection;
import java.net.URL;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.*;

/**
 * @Description: 节假日处理
 * @Author: wangruitao
 * @DATE: 2019/12/2
 **/

public class MetadataUtils {
    private static final Logger log = LoggerFactory.getLogger(MetadataUtils.class);
    public static List<HolidayRecord> readHolidayExcel(String filePath, String editor)throws Exception{

        if(StringUtils.isBlank(filePath)){
            log.info("文件路径不存在！");
            return null;
        }
        InputStream in = new FileInputStream(filePath);//excel文件
        List<HolidayRecord> holidays = getHolidayRecordsByStream(editor, in);

        return holidays;
    }

    public static List<HolidayRecord> readHolidayExcelByUrl(String fileUrl,String editor)throws Exception{

        if(StringUtils.isBlank(fileUrl)){
            log.info("fileUrl不存在！");
            return null;
        }
        URL url = new URL(fileUrl);
        HttpURLConnection connection = (HttpURLConnection) url.openConnection();
        BufferedInputStream in = new BufferedInputStream(connection.getInputStream());

        List<HolidayRecord> holidays = getHolidayRecordsByStream(editor, in);


        return holidays;
    }

    private static List<HolidayRecord> getHolidayRecordsByStream(String editor, InputStream in) throws Exception {
        Workbook book = ImportExcelUtil.getWorkBook(in);
        List<List<String>> list = ImportExcelUtil.getBankStringListByExcel(book);
        in.close();

        DateFormat format1 = new SimpleDateFormat("yyyy-MM-dd");
        Date now = new Date();
        //将数据转换为List<HolidayRecord>类型
        List<HolidayRecord> holidays = new ArrayList<>();
        if (list != null && list.size() > 0) {
            for (int i = 0; i < list.size(); i++) {
                List<String> lo = list.get(i);
                if(StringUtils.isBlank(lo.get(1))){
                    continue;
                }

                HolidayRecord holiday = new HolidayRecord();

//                holiday.setId(Long.valueOf(lo.get(0)));

                Calendar calendar = new GregorianCalendar(1900, 0, -1);
                Date start = calendar.getTime();
                Date d = DateUtils.addDays(start, Integer.valueOf(lo.get(1)));
                holiday.setLocalDate(d);

                holiday.setHolidayType(lo.get(2));
                holiday.setCurrency(lo.get(3));
                holiday.setTimeRange(lo.get(4));
                holiday.setEditor(editor);
                holiday.setGmtCreate(now);
                holiday.setGmtUpdate(now);
                holidays.add(holiday);
            }
        }
        return holidays;
    }


}

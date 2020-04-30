package com.xuxueli.poi.excel.test;

import com.xuxueli.poi.excel.ExcelExportUtil;
import com.xuxueli.poi.excel.ExcelImportUtil;
import com.xuxueli.poi.excel.MyExcelImportUtil;
import com.xuxueli.poi.excel.test.model.HolidayRecord;
import com.xuxueli.poi.excel.test.model.ShopDTO;

import java.util.ArrayList;
import java.util.Date;
import java.util.List;

/**
 * FUN Test
 *
 * @author xuxueli 2017-09-08 22:41:19
 */
public class Test {

    public static void main(String[] args) {
//        shopTest();
        /*List<HolidayRecord> holidayDTOList = new ArrayList<HolidayRecord>();
        for (int i = 0; i < 100; i++) {
            HolidayRecord holiday = new HolidayRecord();

            holidayDTOList.add(holiday);
        }*/
        String filePath = "E:\\节假日记录.xlsx";

        /**
         * Excel导出：Object 转换为 Excel
         */
//        ExcelExportUtil.exportToFile(filePath, holidayDTOList);

        /**
         * Excel导入：Excel 转换为 Object
         */
        List<Object> list = MyExcelImportUtil.importExcel(filePath, HolidayRecord.class);

        System.out.println(list);
    }

    private static void shopTest() {
        /**
         * Mock数据，Java对象列表
         */
        List<ShopDTO> shopDTOList = new ArrayList<ShopDTO>();
        for (int i = 0; i < 100; i++) {
            ShopDTO shop = new ShopDTO(true, "商户"+i, (short) i, 1000+i, 10000+i, (float) (1000+i), (double) (10000+i), new Date());
            shopDTOList.add(shop);
        }
        String filePath = "/Users/xuxueli/Downloads/demo-sheet.xls";

        /**
         * Excel导出：Object 转换为 Excel
         */
        ExcelExportUtil.exportToFile(filePath, shopDTOList);

        /**
         * Excel导入：Excel 转换为 Object
          */
        List<Object> list = ExcelImportUtil.importExcel(filePath, ShopDTO.class);

        System.out.println(list);
    }

}

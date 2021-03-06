package com.xuxueli.poi.excel;

import com.xuxueli.poi.excel.annotation.ExcelField;
import com.xuxueli.poi.excel.annotation.ExcelSheet;
import com.xuxueli.poi.excel.util.FieldReflectionUtil;
import org.apache.commons.lang3.math.NumberUtils;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.formula.functions.T;
import org.apache.poi.ss.usermodel.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.lang.reflect.Modifier;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

/**
 * Excel导入工具
 *
 * @author xuxueli 2017-09-08 22:41:19
 */
public class MyExcelImportUtil {
    private static Logger logger = LoggerFactory.getLogger(MyExcelImportUtil.class);

    /**
     * 从Workbook导入Excel文件，并封装成对象
     *
     * @param workbook
     * @param sheetClass
     * @return List<Object>
     */
    public static List<Object> importExcel(Workbook workbook, Class<?> sheetClass) {
        List<Object> sheetDataList = importSheet(workbook, sheetClass);
        return sheetDataList;
    }

    public static List<Object> importSheet(Workbook workbook, Class<?> sheetClass) {
        try {
            List<Object> dataList = new ArrayList<Object>();
            // sheet
            ExcelSheet excelSheet = sheetClass.getAnnotation(ExcelSheet.class);
            String sheetName = (excelSheet!=null && excelSheet.name()!=null && excelSheet.name().trim().length()>0)?excelSheet.name().trim():sheetClass.getSimpleName();

            // sheet field
            List<Field> fields = new ArrayList<Field>();
            if (sheetClass.getDeclaredFields()!=null && sheetClass.getDeclaredFields().length>0) {
                for (Field field: sheetClass.getDeclaredFields()) {
                    if (Modifier.isStatic(field.getModifiers())) {
                        continue;
                    }
                    fields.add(field);
                }
            }

            if (fields==null || fields.size()==0) {
                throw new RuntimeException(">>>>>>>>>>> xxl-excel error, data field can not be empty.");
            }

            //获取指定工作表<这里获取的是第一个>
            Sheet s = workbook.getSheetAt(NumberUtils.INTEGER_ZERO);
            //循环行sheet.getPhysicalNumberOfRows()是获取表格的总行数
            for (int i = NumberUtils.INTEGER_ONE; i < s.getPhysicalNumberOfRows(); i++) {
                Row row = s.getRow(i);
                Object t =  sheetClass.newInstance();
                // 取出第i行  getRow(index) 获取第(index+1)行
                for (int j = NumberUtils.INTEGER_ZERO; j < row.getLastCellNum(); j++) {
                    // getPhysicalNumberOfCells() 获取当前行的总列数
                    String key = getCellFormatValue(s.getRow(NumberUtils.INTEGER_ZERO).getCell(j)).trim();
                    for (Field field : sheetClass.getDeclaredFields()) {
                        if (field.getAnnotation(ExcelField.class) == null){
                            System.out.println("------------------");
                            continue;
                        }
                        String value =  field.getAnnotation(ExcelField.class).name();
                        if (key.equals(value)) {
//                            field.setAccessible(true);
//                            field.set(t, getCellFormatValue(row.getCell(j)));

//                            String fieldValueStr = row.getCell(j).getStringCellValue();

                            Cell cell = row.getCell(j);
                            cell.setCellType(CellType.STRING);//numberformat异常，不加

                            String fieldValueStr = cell.getStringCellValue();
                            Object fieldValue = FieldReflectionUtil.parseValue(field, fieldValueStr);
                            if (fieldValue == null) {
                                continue;
                            }

                            // fill val
                            field.setAccessible(true);
                            field.set(t, fieldValue);

                            break;
                        }
                    }
                }
                dataList.add(t);
            }



            // sheet data
            /*Sheet sheet = workbook.getSheet(sheetName);
            if (sheet == null) {
                return null;
            }

            Iterator<Row> sheetIterator = sheet.rowIterator();
            int rowIndex = 0;

            while (sheetIterator.hasNext()) {
                Row rowX = sheetIterator.next();
                if (rowIndex > 0) {
                    Object rowObj = sheetClass.newInstance();


                    for (int i = 0; i < fields.size(); i++) {

                        // cell
                        Cell cell = rowX.getCell(i);

                        if (cell == null) {
                            continue;
                        }

                        //test:
                        System.out.println(cell);

//                        continue;

                        // call val str
                        cell.setCellType(CellType.STRING);
                        String fieldValueStr = cell.getStringCellValue();       // cell.getCellTypeEnum()

                        // java val
                        Field field = fields.get(i);
                        Object fieldValue = FieldReflectionUtil.parseValue(field, fieldValueStr);
                        if (fieldValue == null) {
                            continue;
                        }

                        // fill val
                        field.setAccessible(true);
                        field.set(rowObj, fieldValue);
                    }
                    dataList.add(rowObj);
                }
                rowIndex++;
            }*/
            return dataList;
        } catch (IllegalAccessException e) {
            logger.error(e.getMessage(), e);
            throw new RuntimeException(e);
        } catch (InstantiationException e) {
            logger.error(e.getMessage(), e);
            throw new RuntimeException(e);
        }
    }

    /**
     * 导入Excel文件，并封装成对象
     *
     * @param excelFile
     * @param sheetClass
     * @return List<Object>
     */
    public static List<Object> importExcel(File excelFile, Class<?> sheetClass) {
        try {
            Workbook workbook = WorkbookFactory.create(excelFile);
            List<Object> dataList = importExcel(workbook, sheetClass);
            return dataList;
        } catch (IOException e) {
            logger.error(e.getMessage(), e);
            throw new RuntimeException(e);
        }
    }

    /**
     * 从文件路径导入Excel文件，并封装成对象
     *
     * @param filePath
     * @param sheetClass
     * @return List<Object>
     */
    public static List<Object> importExcel(String filePath, Class<?> sheetClass) {
        File excelFile = new File(filePath);
        List<Object> dataList = importExcel(excelFile, sheetClass);
        return dataList;
    }

    /**
     * 导入Excel数据流，并封装成对象
     *
     * @param inputStream
     * @param sheetClass
     * @return List<Object>
     */
    public static List<Object> importExcel(InputStream inputStream, Class<?> sheetClass) {
        try {
            Workbook workbook = WorkbookFactory.create(inputStream);
            List<Object> dataList = importExcel(workbook, sheetClass);
            return dataList;
        } catch (IOException e) {
            logger.error(e.getMessage(), e);
            throw new RuntimeException(e);
        }
    }


    public static String getCellFormatValue(Cell cell) {
        String cellValue = "";
        if (cell != null) {
            //判断cell类型
            switch (cell.getCellType()) {
                case NUMERIC: {
                    cellValue = String.valueOf(cell.getNumericCellValue());
                    break;
                }
                case STRING: {
                    cellValue = cell.getRichStringCellValue().getString();
                    break;
                }
                case BOOLEAN: {
                    cellValue = String.valueOf(cell.getBooleanCellValue());
                    break;
                }
                default:
                    cellValue = "";
            }
        }
        return cellValue;
    }

}

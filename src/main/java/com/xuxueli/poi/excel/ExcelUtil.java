package com.xuxueli.poi.excel;

import com.alibaba.fastjson.JSONArray;
import com.fasterxml.jackson.annotation.JsonProperty;

import org.apache.commons.lang3.StringUtils;
import org.apache.commons.lang3.math.NumberUtils;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.OfficeXmlFileException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.*;
import java.lang.reflect.Field;
import java.util.*;

/**
 * @author：shihuiwen
 * @date：Created in 2019/12/6 13:24
 * @description：处理excel
 * @modified By：
 * @version:1.0$
 */
public class ExcelUtil {

    private static final Logger LOGGER = LoggerFactory.getLogger(ExcelUtil.class);

    /**
     * 当前目录路径
     */
    private static final String CURRENTWORKDIR = System.getProperty("user.dir") + "/file/";

    private ExcelUtil() {
    }

    /**
     * 生成excel
     *
     * @param report
     * @return
     */
    public static String generateExcel(String report, String reportId) throws Exception {
        if (StringUtils.isBlank(report)) {
            throw new Exception("excel生成失败");
        }
        try {
            String indexString;
            int sub;
            if(report.contains("\r\n")){
                indexString = "\r\n";
                sub = 2;
            }else{
                indexString = "\n";
                sub = 1;
            }
            String title = report.substring(0, report.indexOf(indexString));
            List<String> titleList = Arrays.asList(title.split("\t"));
            String data = report.substring(report.indexOf(indexString) + sub);
            List<String> dataStringList = Arrays.asList(data.split(indexString));
            List<Map<String, String>> dataListMap = new ArrayList<>();
            for (String dataString : dataStringList) {
                Map<String, String> dataMap = new TreeMap<>();
                List<String> dataList = Arrays.asList(dataString.split("\t"));
                List<String> newDataList = new ArrayList<>(dataList);
                newDataList.add("");
                newDataList.add("");
                for (int i = 0; i < titleList.size(); i++) {
                    dataMap.put(titleList.get(i), newDataList.get(i));
                }
                dataListMap.add(dataMap);
            }

            JSONArray datajson = (JSONArray) JSONArray.toJSON(dataListMap);
            generateExcel(CURRENTWORKDIR, reportId+".xlsx", titleList, titleList, datajson);
        } catch (Exception e) {
            LOGGER.error("生成excel文件失败",e);
            return null;
        }
        return CURRENTWORKDIR+reportId+".xlsx";
    }

    public static boolean generateExcel(String filepath, String filename, List<String> titlelist, List<String> zdlist, JSONArray datalist) throws IOException{
        boolean success = false;
        try {
            //创建HSSFWorkbook对象(excel的文档对象)
            HSSFWorkbook wb = new HSSFWorkbook();
            // 建立新的sheet对象（excel的表单）
            HSSFSheet sheet = wb.createSheet("sheet1");
            // 在sheet里创建第一行，参数为行索引(excel的行)，可以是0～65535之间的任何一个
            HSSFRow row0 = sheet.createRow(0);
            // 添加表头
            for(int i = 0;i<titlelist.size();i++){
                row0.createCell(i).setCellValue(titlelist.get(i));
            }
            //添加表中内容
            for(int row = 0;row<datalist.size();row++){//数据行
                //创建新行
                HSSFRow newrow = sheet.createRow(row+1);//数据从第二行开始
                //获取该行的数据
                @SuppressWarnings("unchecked")
                Map<String,Object> data = (Map<String, Object>) datalist.get(row);

                for(int col = 0;col<zdlist.size();col++){//列
                    //数据从第一列开始
                    //创建单元格并放入数据
                    newrow.createCell(col).setCellValue(data!=null&&data.get(zdlist.get(col))!=null?String.valueOf(data.get(zdlist.get(col))):"");
                }
            }

            //判断是否存在目录. 不存在则创建
            isChartPathExist(filepath);
            //输出Excel文件1
            FileOutputStream output=new FileOutputStream(filepath+filename);
            wb.write(output);//写入磁盘
            output.close();
            success = true;
        } catch (Exception e) {
            success = false;
            LOGGER.error("", e);
        }
        return success;
    }


    /**
     * 判断文件夹是否存在，如果不存在则新建
     *
     * @param dirPath 文件夹路径
     */
    private static void isChartPathExist(String dirPath) {
        File file = new File(dirPath);
        if (!file.exists()) {
            file.mkdirs();
        }
    }

    /**
     * 解析excel
     *
     * @param clazz
     * @param fileName
     * @param <T>
     * @return
     */
    public static <T> List<T> resolve(Class<?> clazz, String fileName) throws Exception {
        if (clazz == null || StringUtils.isEmpty(fileName)) {
            throw new Exception("ResultCode.PARAM_IS_BLANK");
        }
        List<T> response = new ArrayList<>();
        //获得excel文件对象workbook
        Workbook wb;
        try {
            wb = readExcel(fileName);
            if (wb == null) {
                throw new Exception( "读取excel文件流失败");
            }
        } catch (OfficeXmlFileException e) {
            try {
                LOGGER.info("更换解析文件XSSF");
                InputStream is = new FileInputStream(fileName);
                wb = new XSSFWorkbook(is);
            } catch (FileNotFoundException ex) {
                LOGGER.error("", e);
                throw new Exception(e.getMessage());
            } catch (IOException ex) {
                LOGGER.error("", e);
                throw new Exception(e.getMessage());
            }
        }
        try {
            //获取指定工作表<这里获取的是第一个>
            Sheet s = wb.getSheetAt(NumberUtils.INTEGER_ZERO);
            //循环行sheet.getPhysicalNumberOfRows()是获取表格的总行数
            for (int i = NumberUtils.INTEGER_ONE; i < s.getPhysicalNumberOfRows(); i++) {
                Row row = s.getRow(i);
                T t = (T) clazz.newInstance();
                // 取出第i行  getRow(index) 获取第(index+1)行
                for (int j = NumberUtils.INTEGER_ZERO; j < row.getLastCellNum(); j++) {
                    // getPhysicalNumberOfCells() 获取当前行的总列数
                    String key = getCellFormatValue(s.getRow(NumberUtils.INTEGER_ZERO).getCell(j)).trim();
                    for (Field field : clazz.getDeclaredFields()) {
                        String value = SafeUtils.safeGet(() -> field.getAnnotation(JsonProperty.class).value());
                        if (key.equals(value)) {
                            field.set(t, getCellFormatValue(row.getCell(j)));
                            break;
                        }
                    }
                }
                response.add(t);
            }
        } catch (IndexOutOfBoundsException e) {
            LOGGER.warn("", e);
            throw new Exception( e.getMessage());
        } catch (IllegalAccessException e) {
            LOGGER.warn("", e);
            throw new Exception(e.getMessage());
        } catch (InstantiationException e) {
            LOGGER.warn("", e);
            throw new Exception( e.getMessage());
        }
        return response;
    }

    /**
     * xls/xlsx都使用的Workbook
     *
     * @param fileName
     * @return
     */
    public static HSSFWorkbook readExcel(String fileName) throws Exception {
        try {
            InputStream is = new FileInputStream(fileName);
            return new HSSFWorkbook(is);
        } catch (FileNotFoundException e) {
            LOGGER.error("", e);
            throw new Exception( e.getMessage());
        } catch (IOException e) {
            LOGGER.error("", e);
            throw new Exception(e.getMessage());
        }
    }

    /**
     * format表格内容
     *
     * @param cell
     * @return
     */
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

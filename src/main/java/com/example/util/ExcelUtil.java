package com.example.exceldemo.util;


import lombok.extern.slf4j.Slf4j;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Component;


import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;


/***
 * @description: Excel操作
 * @author: zhangting
 * @date: 2021/9/17 7:14 下午
 **/
@Component
@Slf4j
public class ExcelUtil {


    /**
     * 课程excel
     *
     * @param in
     * @param fileName
     * @return
     * @throws Exception
     */
    public static void readExcelContent(InputStream in, String fileName) throws Exception {
        // 创建excel工作簿
        Workbook work = getWorkbook(in, fileName);
        if (null == work) {
            throw new Exception("创建Excel工作薄为空！");
        }
        Sheet sheet = null;
        Row row = null;
        Cell cell = null;
        for (int i = 0; i < work.getNumberOfSheets(); i++) {
            sheet = work.getSheetAt(i);
            if (sheet == null) {
                continue;
            }
            for (int j = 0; j <= sheet.getLastRowNum(); j++) {
                row = sheet.getRow(j);
                if (row == null) {
                    continue;
                }
                for (int y = row.getFirstCellNum(); y < row.getLastCellNum(); y++) {
                    cell = row.getCell(y);
                    log.info("rowIndex:{},columnIndex:{},cellValue:{}", j, y, getCellValue(cell));
                }
            }
        }
        work.close();
        return;
    }

    /**
     * 判断文件格式
     *
     * @param in
     * @param fileName
     * @return
     */
    private static Workbook getWorkbook(InputStream in, String fileName) throws Exception {
        Workbook book = null;
        String filetype = fileName.substring(fileName.lastIndexOf("."));
        if (".xls".equals(filetype)) {
            book = new HSSFWorkbook(in);
        } else if (".xlsx".equals(filetype)) {
            book = new XSSFWorkbook(in);
        } else {
            throw new Exception("请上传excel文件！");
        }

        return book;
    }


    /**
     * 读取Excel
     */
    public static void readExcel() {
        File file = new File("/Users/admin/Desktop/模板.xlsx");
        try {
            InputStream inputStream = new FileInputStream(file);
            readExcelContent(inputStream, file.getName());
            inputStream.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public static void main(String[] args) {
        readExcel();
    }


    /**
     * 读取Excel单元格内容
     *
     * @param cell
     * @return
     */
    private static String getCellValue(Cell cell) {
        String cellValue = "";
        switch (cell.getCellType()) {
            case STRING:
                cellValue = cell.getStringCellValue().trim();
                break;
            case NUMERIC:
                cellValue = String.valueOf(cell.getNumericCellValue());
                break;
            case BOOLEAN:
                cellValue = String.valueOf(cell.getBooleanCellValue());
                break;
            default:
                cellValue = "";
        }
        return cellValue;
    }

}

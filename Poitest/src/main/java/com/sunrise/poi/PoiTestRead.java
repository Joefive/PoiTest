package com.sunrise.poi;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.joda.time.DateTime;
import org.junit.Test;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Date;

public class PoiTestRead {
    String PATH = "G:\\Down\\";
    FileInputStream fileInputStream = null;

    @Test
    public void testRead01() throws IOException {
        try {
            //1.获取文件流
            fileInputStream = new FileInputStream(PATH + "ReadExcel.xlsx");
            //2.获取workbook对象工作簿
            Workbook workbook = new XSSFWorkbook(fileInputStream);
            //3.获取工作表sheet
            Sheet sheet = workbook.getSheetAt(0);
            //4.获取行
            Row row = sheet.getRow(0);
            //5.获取列
            Cell cell = row.getCell(0);
            //打印获取单元格值
            String stringCellValue = cell.getStringCellValue();
            System.out.println(stringCellValue);

        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } finally {
            fileInputStream.close();
        }


    }

    @Test
    public void testRead02() throws IOException {
        try {
            //1.获取文件流
            fileInputStream = new FileInputStream(PATH + "ReadExcel.xlsx");
            //2.获取workbook对象工作簿
            Workbook workbook = new XSSFWorkbook(fileInputStream);
            //3.获取工作表sheet
            Sheet sheet = workbook.getSheetAt(0);
            //4.获取首行(标题行)
            Row rowTitle = sheet.getRow(0);
            if (rowTitle != null) {
                int cellCount = rowTitle.getPhysicalNumberOfCells();
                for (int cellNum = 0; cellNum < cellCount; cellNum++) {
                    Cell cell = rowTitle.getCell(cellNum);
                    if (cell != null) {
                        int cellType = cell.getCellType();
                        String stringCellValue = cell.getStringCellValue();
                        System.out.print(stringCellValue + "|");
                    }
                }
            }
            System.out.println();
            //5.获取总行数
            int rowCount = sheet.getPhysicalNumberOfRows();
            //6.循环每一行,除过标题行所有行数加1从第1行开始
            for (int rowNum = 1; rowNum < rowCount; rowNum++) {
                Row rowDate = sheet.getRow(rowNum);
                if (rowDate != null) {
                    //7.读取列数据，读取列总数
                    int cellCount = rowDate.getPhysicalNumberOfCells();
                    //8.循环取出列的值
                    for (int cellNum = 0; cellNum < cellCount; cellNum++) {
                        System.out.print("【" + (rowNum + 1) + "-" + cellNum + "】");
                        Cell cell = rowDate.getCell(cellNum);
                        //9.匹配每列数据类型
                        if (cell != null) {
                            int cellType = cell.getCellType();
                            String cellValue = "";
                            switch (cellType) {
                                case HSSFCell.CELL_TYPE_STRING: //判断是否字符串
                                    System.out.print("String:");
                                    cellValue = cell.getStringCellValue();
                                    break;
                                case HSSFCell.CELL_TYPE_BOOLEAN: //判断是否为布尔
                                    System.out.print("Boolean:");
                                    cellValue = String.valueOf(cell.getBooleanCellValue());
                                    break;
                                case HSSFCell.CELL_TYPE_BLANK: //判断是否为空
                                    System.out.print("Null");
                                    break;
                                case HSSFCell.CELL_TYPE_NUMERIC: //判断是否为数字类型，需要二次判断是否为数字&日期
                                    System.out.print("Numeric:");
                                    if (HSSFDateUtil.isCellInternalDateFormatted(cell)) {
                                        System.out.print("日期：");
                                        Date dateCellValue = cell.getDateCellValue();
                                        cellValue = new DateTime(dateCellValue).toString("yyyy-MM-dd");
                                    } else {
                                        //如果不是日期格式，防止数字太长直接转换成字符串输出
                                        System.out.print("转换字符串：");
                                        //cell.setCellValue(HSSFCell.CELL_TYPE_STRING);
                                        //cellValue = cell.toString();
                                        double numericCellValue = cell.getNumericCellValue();
                                        cellValue = String.valueOf(numericCellValue);
                                    }
                                    break;
                                case HSSFCell.CELL_TYPE_ERROR:
                                    System.out.print("类型错误！");
                                    break;
                            }
                            System.out.print(cellValue);
                        }
                    }
                    System.out.println();
                }

            }
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } finally {
            fileInputStream.close();
        }


    }
}

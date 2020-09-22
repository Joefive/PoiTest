package com.sunrise.poi;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.joda.time.DateTime;
import org.junit.Test;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

public class PoiTestWrite {
    FileOutputStream fileOutputStream = null;
    String PATH = "G:\\Down\\";

    /**
     * 03版本测试，HSSFWorkbook对象
     */
    @Test
    public void testWrite03() {
        //创建工作簿
        Workbook workbook = new HSSFWorkbook();
        //创建工作表SHEET
        Sheet sheet = workbook.createSheet("sx_count");
        //创建第一行1,1
        Row row1 = sheet.createRow(0);
        //创建单元并赋值1,1
        Cell cell = row1.createCell(0);
        cell.setCellValue("AREA_CODE");
        //赋值1,2
        Cell cell1 = row1.createCell(1);
        cell1.setCellValue("AREA_NAME");
        Cell create_time = row1.createCell(2);
        create_time.setCellValue("CREATE_TIME");

        //创建第二行，2,1
        Row row2 = sheet.createRow(1);
        Cell cell2 = row2.createCell(0);
        cell2.setCellValue("610000");
        //赋值2,2
        Cell cell3 = row2.createCell(1);
        cell3.setCellValue("陕西省");
        //赋值2,3
        Cell cell4 = row2.createCell(2);
        String time = new DateTime().toString("yyyy-MM-dd HH:mm:ss");
        cell4.setCellValue(time);

        try {
            fileOutputStream = new FileOutputStream(PATH + "SX_COUNT_2003Excel.xls");
            try {
                workbook.write(fileOutputStream);
                System.out.println("文件：" + PATH + "SX_COUNT_2003Excel.xls" + ",已生成！");
            } catch (IOException e) {
                e.printStackTrace();
            }
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } finally {
            try {
                fileOutputStream.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }

    /**
     * 07版本测试，XSSFWorkbook对象
     */
    @Test
    public void testWrite07() {
        //创建工作簿
        Workbook workbook = new XSSFWorkbook();
        //创建工作表SHEET
        Sheet sheet = workbook.createSheet("sx_count");
        //创建第一行1,1
        Row row1 = sheet.createRow(0);
        //创建单元并赋值1,1
        Cell cell = row1.createCell(0);
        cell.setCellValue("AREA_CODE");
        //赋值1,2
        Cell cell1 = row1.createCell(1);
        cell1.setCellValue("AREA_NAME");
        Cell create_time = row1.createCell(2);
        create_time.setCellValue("CREATE_TIME");

        //创建第二行，2,1
        Row row2 = sheet.createRow(1);
        Cell cell2 = row2.createCell(0);
        cell2.setCellValue("610000");
        //赋值2,2
        Cell cell3 = row2.createCell(1);
        cell3.setCellValue("陕西省");
        //赋值2,3
        Cell cell4 = row2.createCell(2);
        String time = new DateTime().toString("yyyy-MM-dd HH:mm:ss");
        cell4.setCellValue(time);

        try {
            fileOutputStream = new FileOutputStream(PATH + "SX_COUNT_2007Excel.xlsx");
            try {
                workbook.write(fileOutputStream);
                System.out.println("文件：" + PATH + "SX_COUNT_2007Excel.xlsx" + ",已生成！");
            } catch (IOException e) {
                e.printStackTrace();
            }
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } finally {
            try {
                fileOutputStream.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }

    /**
     * HSSFWorkbook为03版对象可以创建65536行数据，生成结果快
     * XSSFWorkbook为07版本对象可以创建无数行数据，但是生成结果较慢
     * 解决方法：大文件可以进行设置XSSFWorkbook对象的条数，默认100条记录在内存中，超过则写入到临时文件中
     */
    @Test
    public void testBigData() throws IOException {
        long begin = System.currentTimeMillis();
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet();
        for (int rowNum = 0; rowNum < 65537; rowNum++) {
            Row row = sheet.createRow(rowNum);
            for (int cellNum = 0; cellNum < 10; cellNum++) {
                Cell cell = row.createCell(cellNum);
                cell.setCellValue(cellNum);
            }
        }
        fileOutputStream = new FileOutputStream(PATH + "testBigData.xlsx");
        workbook.write(fileOutputStream);
        fileOutputStream.close();
        System.out.println("写入完毕...!");
        long end = System.currentTimeMillis();
        System.out.println("写入时间:" + (double) (end - begin) / 1000 + "S");
    }

    /**
     * 测试SXSSFWorkbook对象，使用临时文件方法
     *
     * @throws IOException
     */
    @Test
    public void testBigDataSuper() throws IOException {
        long begin = System.currentTimeMillis();
        Workbook workbook = new SXSSFWorkbook();
        Sheet sheet = workbook.createSheet();
        for (int rowNum = 0; rowNum < 100000; rowNum++) {
            Row row = sheet.createRow(rowNum);
            for (int cellNum = 0; cellNum < 10; cellNum++) {
                Cell cell = row.createCell(cellNum);
                cell.setCellValue(cellNum);
            }
        }
        fileOutputStream = new FileOutputStream(PATH + "testBigDataSuper.xlsx");
        workbook.write(fileOutputStream);
        fileOutputStream.close();
        System.out.println("写入完毕...!");
        long end = System.currentTimeMillis();
        //每次生成100条数据后需要关闭对象，清除临时文件
        ((SXSSFWorkbook) workbook).dispose();
        System.out.println("写入时间:" + (double) (end - begin) / 1000 + "S");
    }
}

package com.sunrise.poi;

import com.alibaba.druid.pool.DruidDataSourceFactory;

import javax.sql.DataSource;
import java.io.InputStream;
import java.sql.Connection;
import java.util.Properties;

public class DataSourceDruid {
    public static void main(String[] args) throws Exception {
        Properties pro = new Properties();
        String PATH = "E:\\GITHUB\\PoiTest\\Poitest\\src\\druid.properties";
        InputStream is = DataSourceDruid.class.getClassLoader().getResourceAsStream(PATH);
        pro.load(is);
        //4.创建连接池对象
        DataSource ds = DruidDataSourceFactory.createDataSource(pro);
        for (int i = 0; i < 10; i++) {
            //5.获取连接池对象
            Connection conn = ds.getConnection();
            System.out.println(i + ":" + conn);
            if (i == 5) {
                conn.close();
            }
        }

    }
}

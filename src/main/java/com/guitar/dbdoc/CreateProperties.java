package com.guitar.dbdoc;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Properties;

/**
 *
 * @author hxy
 */
public class CreateProperties {
    public static void create(File file) throws IOException{
        Properties prop = new Properties();
        prop.setProperty("url","jdbc:mysql://127.0.0.1:3306/");
        prop.setProperty("db","mysql");
        prop.setProperty("user","root");
        prop.setProperty("pwd","toor");
        
        prop.store(new FileOutputStream(file), "数据连接配置文件");
    }
}

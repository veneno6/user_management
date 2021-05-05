package com.veo;

import com.opencsv.CSVReader;
import com.veo.pojo.User;

import java.io.FileReader;
import java.text.SimpleDateFormat;

//使用openCSV导入百万数据
public class CSVImportMillionData {

    private static SimpleDateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd");

    public static void main(String[] args) throws Exception {
        //使用csvReader对象读取Excel文件,文件的绝对路径，需前面导出文件后测试
        CSVReader csvReader = new CSVReader(new FileReader("d:\\百万用户数据的导出.csv"));
        //先读取标题
        String[] title = csvReader.readNext();
        //如果使用readAll方法，或报内存溢出的异常，要逐行读取文件
        User user = null;
        while(true){
            String[] content = csvReader.readNext();
            user = new User();
            if (content == null){
                break;
            }
            user.setId(Long.parseLong(content[0]));
            user.setUserName(content[1]);
            user.setPhone(content[2]);
            user.setBirthday(dateFormat.parse(content[3]));
            user.setAddress(content[4]);
            System.out.println(user);
        }
    }
}

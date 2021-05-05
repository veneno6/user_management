package com.veo;

import com.veo.pojo.User;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

import java.text.SimpleDateFormat;
import java.util.Date;

public class POIReadExcelFile {

    private SimpleDateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd");

    //使用POI读取，导入Excel中的内容
    @Test
    public void ReadExcelFile() throws Exception {
        //创建一个可读取文件的工作薄
        Workbook workbook = new XSSFWorkbook("用户导入测试数据.xlsx");
        //获取第一个工作表
        Sheet sheet = workbook.getSheetAt(0);
        //获取工作表的最后一行的索引值，用于遍历表中要读取的数据
        int lastRowIndex = sheet.getLastRowNum();
        //从第二行开始读取数据，跳过标题行，并等于最后一行的索引
        Row row = null;
        User user = null;
        for (int i = 1; i <= lastRowIndex ; i++) {
//  标题列：用户名 	手机号	省份	城市	工资	入职日期	出生日期	现住地址
            //每次获取一行
            row = sheet.getRow(i);
            //获取行中的单元的内容，并在数据类型的转换
            String username = row.getCell(0).getStringCellValue();
            //手机号码在用户输入的时候可能为数字类型，做一个异常的处理
            String phone = null;
            try {
                phone = row.getCell(1).getStringCellValue();
            } catch (Exception e) {
                phone = row.getCell(1).getNumericCellValue()+"";
            }
            String province = row.getCell(2).getStringCellValue();
            String city = row.getCell(3).getStringCellValue();
            //处理数值类型,double类型转int
            int salary = ((Double)row.getCell(4).getNumericCellValue()).intValue();
            //使用SimpleDateFormat,处理日期类型
            Date hireDate = dateFormat.parse(row.getCell(5).getStringCellValue());
            Date birthDay = dateFormat.parse(row.getCell(6).getStringCellValue());
            String address = row.getCell(7).getStringCellValue();

            user = new User();
            user.setUserName(username);
            user.setPhone(phone);
            user.setProvince(province);
            user.setCity(city);
            user.setSalary(salary);
            user.setHireDate(hireDate);
            user.setBirthday(birthDay);
            user.setAddress(address);

            System.out.println(user);
        }

    }
}

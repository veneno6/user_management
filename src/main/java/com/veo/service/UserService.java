package com.veo.service;

import cn.afterturn.easypoi.csv.CsvExportUtil;
import cn.afterturn.easypoi.csv.entity.CsvExportParams;
import cn.afterturn.easypoi.entity.ImageEntity;
import cn.afterturn.easypoi.excel.ExcelExportUtil;
import cn.afterturn.easypoi.excel.ExcelImportUtil;
import cn.afterturn.easypoi.excel.entity.ExportParams;
import cn.afterturn.easypoi.excel.entity.ImportParams;
import cn.afterturn.easypoi.excel.entity.TemplateExportParams;
import cn.afterturn.easypoi.excel.entity.enmus.ExcelType;
import cn.afterturn.easypoi.word.WordExportUtil;
import com.github.pagehelper.Page;
import com.github.pagehelper.PageHelper;
import com.opencsv.CSVWriter;
import com.veo.mapper.ResourceMapper;
import com.veo.mapper.UserMapper;
import com.veo.pojo.Resource;
import com.veo.pojo.User;
//import jxl.Workbook;
//import org.apache.poi.ss.usermodel.Workbook;
import com.veo.utils.EntityUtils;
import com.veo.utils.ExcelExportEngine;
import com.zaxxer.hikari.HikariDataSource;
import jxl.Workbook;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import net.sf.jasperreports.engine.JREmptyDataSource;
import net.sf.jasperreports.engine.JasperExportManager;
import net.sf.jasperreports.engine.JasperFillManager;
import net.sf.jasperreports.engine.JasperPrint;
import net.sf.jasperreports.engine.data.JRBeanCollectionDataSource;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.util.Units;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.xwpf.usermodel.*;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.util.CollectionUtils;
import org.springframework.util.ResourceUtils;
import org.springframework.web.multipart.MultipartFile;
import tk.mybatis.mapper.entity.Example;


import javax.imageio.ImageIO;
import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletResponse;
import java.awt.image.BufferedImage;
import java.io.*;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.stream.Collectors;

@Service
public class UserService {

    @Autowired
    private UserMapper userMapper;

    @Autowired
    private ResourceMapper resourceMapper;

    @Autowired
    private HikariDataSource hikariDataSource;

    private SimpleDateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd");

    public List<User> findAll() {
        return userMapper.selectAll();
    }

    public List<User> findPage(Integer page, Integer pageSize) {
        //开启分页
        PageHelper.startPage(page,pageSize);
        //实现查询
        Page<User> userPage = (Page<User>) userMapper.selectAll();
        return userPage.getResult();
    }

    public User findbyId(Long id) {
        //更加id查询用户，并返回带公共用品的数据
        User user = userMapper.selectByPrimaryKey(id);
        //查询公共用品的数据
        Resource resource = new Resource();
        resource.setUserId(user.getId());
        List<Resource> resourceList = resourceMapper.select(resource);
        user.setResourceList(resourceList);
        return user;
    }

    public void downLoadXlsByJxl(HttpServletResponse response) throws Exception {
        // 导出标题： 编号 姓名  手机号 入职日期    现住址

        //使用HttpServletResponse，获取输出流
        ServletOutputStream outputStream = response.getOutputStream();
        //1.创建一个工作薄
        WritableWorkbook workbook = Workbook.createWorkbook(outputStream);

        //2.在创建的工作薄中创建工作表，参数：String 工作表名, int 表索引
        WritableSheet sheet = workbook.createSheet("JXL入门",0);

        //设置单元格的列宽
        //int 列标, int 几个字符的宽度（数字表示）:数字为：width * 256
        sheet.setColumnView(0,5);
        sheet.setColumnView(1,8);
        sheet.setColumnView(2,15);
        sheet.setColumnView(3,15);
        sheet.setColumnView(4,28);

        //3.创建单元格，遍历标题数组，设置固定的表格标题
        Label label =null;
        String[] title = {"编号","姓名","手机号","入职日期","现住址"};
        for (int i = 0; i < title.length; i++) {
            //参数：int 列标, int 行标, String 单元格内容
            //标题在第一行，行标不变，列标变
            label = new Label(i,0,title[i]);
            //4.加入工作表
            sheet.addCell(label);
        }

        //处理表格中的数据部分
        int rowIndex = 1;
        List<User> users = userMapper.selectAll();
        for (User user : users) {
            //列表，行标，编号
            label = new Label(0,rowIndex,user.getId().toString());
            sheet.addCell(label);

            //列表，行标，姓名
            label = new Label(1,rowIndex,user.getUserName());
            sheet.addCell(label);

            //列表，行标，手机号
            label = new Label(2,rowIndex,user.getPhone());
            sheet.addCell(label);

            //列表，行标，入职日期
            label = new Label(3,rowIndex,dateFormat.format(user.getHireDate()));
            sheet.addCell(label);

            //列表，行标，现住址
            label = new Label(4,rowIndex,user.getAddress());
            sheet.addCell(label);
            rowIndex++;
        }

        //导出文件：一个流（OutPutStream），两个头：文件打开方式（in-line,attachment）、文件下载的mime类型
        String fileName = "一个Jxl导出Excel文件.xls";
        response.setHeader("Content-Disposition","attachment;filename="+new String(fileName.getBytes(),"ISO8859-1"));
        response.setContentType("application/vnd.ms-excel");

        //导出文件
        workbook.write();
        //关闭流
        workbook.close();
        outputStream.close();
    }

    public void uploadExcel(MultipartFile file) throws Exception {
        //创建一个可读取文件的工作薄
        XSSFWorkbook workbook = new XSSFWorkbook(file.getInputStream());
        //获取第一个工作表
        Sheet sheet = workbook.getSheetAt(0);
        //获取工作表的最后一行的索引值，用于遍历表中要读取的数据
        int lastRowIndex = sheet.getLastRowNum();
        //从第二行开始读取数据，跳过标题行，并等于最后一行的索引
        Row row = null;
        User user = null;
        for (int i = 1; i <= lastRowIndex; i++) {
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
                phone = row.getCell(1).getNumericCellValue() + "";
            }
            String province = row.getCell(2).getStringCellValue();
            String city = row.getCell(3).getStringCellValue();
            //处理数值类型,double类型转int
            int salary = ((Double) row.getCell(4).getNumericCellValue()).intValue();
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

            //插入用户数据到数据库
            userMapper.insert(user);
        }
    }

    public void downLoadXlsxByPoi(HttpServletResponse response) throws Exception {
        // 导出标题： 编号 姓名  手机号 入职日期    现住址
        //创建工作薄
        XSSFWorkbook workbook = new XSSFWorkbook();
        //创建工作表
        XSSFSheet sheet = workbook.createSheet("用户数据");

        //设置列宽；int 列标, int 几个字符的宽度（数字表示）:数字为：width / 256
        sheet.setColumnWidth(0,5*256);
        sheet.setColumnWidth(1,8*256);
        sheet.setColumnWidth(2,15*256);
        sheet.setColumnWidth(3,15*256);
        sheet.setColumnWidth(4,28*256);

        //处理固定的标题
        Row titleRow = sheet.createRow(0);
        Cell cell = null;
        String[] title = {"编号","姓名","手机号","入职日期","现住址"};
        for (int i = 0; i < title.length; i++) {
            cell = titleRow.createCell(i);
            cell.setCellValue(title[i]);
        }

        //获取用户数据，插入到表格中的第二行
        List<User> users = userMapper.selectAll();
        //第二行开始插入
        int rowIndex = 1;
        Row row = null;
        for (User user : users) {
            // 导出标题： 编号 姓名  手机号 入职日期    现住址
            //每一行数据，为表格中的一行
            row = sheet.createRow(rowIndex);

            cell = row.createCell(0);
            cell.setCellValue(user.getId());

            cell = row.createCell(1);
            cell.setCellValue(user.getUserName());

            cell = row.createCell(2);
            cell.setCellValue(user.getPhone());
            //处理日期的转换
            cell = row.createCell(3);
            cell.setCellValue(dateFormat.format(user.getHireDate()));

            cell = row.createCell(4);
            cell.setCellValue(user.getAddress());
            //行索引加一
            rowIndex++;
        }

        //导出文件：一个流（OutPutStream），两个头：文件打开方式（in-line,attachment）、文件下载的mime类型
        String fileName = "员工数据.xlsx";
        response.setHeader("Content-Disposition","attachment;filename="+new String(fileName.getBytes(),"ISO8859-1"));
        response.setContentType("pplication/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        //导出文件
        workbook.write(response.getOutputStream());
        workbook.close();
    }

    public void downLoadXlsxByPoiWithCellStyle(HttpServletResponse response) throws Exception {
        XSSFWorkbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("用户数据");
        //设置表格的列宽
        sheet.setColumnWidth(0,5*256);
        sheet.setColumnWidth(1,10*256);
        sheet.setColumnWidth(2,15*256);
        sheet.setColumnWidth(3,15*256);
        sheet.setColumnWidth(4,30*256);
        //需求：1.边框线：全边框 2.合并单元格：第一行的第1个单元格到第5个 3.行高：42 4. 对齐方式：水平垂直居中 5.字体:黑体，18
        //使用workbook创建样式，可以用于当前的workbook
        CellStyle bitTitleRowStyle = workbook.createCellStyle();
        //设置样式的上、下、左、右，全边框线，参数：边框线的类型
        bitTitleRowStyle.setBorderTop(BorderStyle.THIN);
        bitTitleRowStyle.setBorderBottom(BorderStyle.THIN);
        bitTitleRowStyle.setBorderLeft(BorderStyle.THIN);
        bitTitleRowStyle.setBorderRight(BorderStyle.THIN);
        //设置对齐的方式：水平、垂直居中对齐,参数：对齐方式
        bitTitleRowStyle.setAlignment(HorizontalAlignment.CENTER);
        bitTitleRowStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        //使用workbook创建字体，可应用于当前的workbook
        Font font = workbook.createFont();
        //字体样式为：黑体 18号
        font.setFontName("黑体");
        font.setFontHeightInPoints((short) 18);
        //应用样式，到标题行
        bitTitleRowStyle.setFont(font);

        //设置大标题，第一行
        Row bitTitleRow = sheet.createRow(0);
        for (int i = 0; i <=4 ; i++) {
            //创建单元格
            Cell cell = bitTitleRow.createCell(i);
            //为单元格应用样式
            cell.setCellStyle(bitTitleRowStyle);
        }
        //合并单元格：int firstRow 起始行, int lastRow 结束行, int firstCol 起始列, int lastCol 结束列
        sheet.addMergedRegion(new CellRangeAddress(0,0,0,4));
        //向创建好的单元格中放入数据
        sheet.getRow(0).getCell(0).setCellValue("用户数据");

        //小标题样式
        CellStyle littleTitleStyle = workbook.createCellStyle();
        //克隆大标题的基本样式，减少代码的冗余
        littleTitleStyle.cloneStyleFrom(bitTitleRowStyle);
        //设置字体：宋体 12号 加粗
        Font littleFont = workbook.createFont();
        littleFont.setFontName("宋体");
        littleFont.setFontHeightInPoints((short) 12);
        littleFont.setBold(true);
        littleTitleStyle.setFont(littleFont);

        //内容样式
        CellStyle contentStyle = workbook.createCellStyle();
        //克隆大标题的基本样式，减少代码的冗余
        contentStyle.cloneStyleFrom(bitTitleRowStyle);
        //水平居左
        contentStyle.setAlignment(HorizontalAlignment.LEFT);
        //设置字体：宋体 11号
        Font contentFont = workbook.createFont();
        contentFont.setFontName("宋体");
        contentFont.setFontHeightInPoints((short) 11);
        contentFont.setBold(false);
        contentStyle.setFont(contentFont);

        //处理固定的标题，第二行
        Row titleRow = sheet.createRow(1);
        Cell cell = null;
        String[] title = {"编号","姓名","手机号","入职日期","现住址"};
        for (int i = 0; i < title.length; i++) {
            cell = titleRow.createCell(i);
            cell.setCellValue(title[i]);
            cell.setCellStyle(littleTitleStyle);
        }

        //获取用户数据，插入到表格中的第二行
        List<User> users = userMapper.selectAll();
        //第三行开始插入
        int rowIndex = 2;
        Row contentRow = null;
        for (User user : users) {
            // 导出标题： 编号 姓名  手机号 入职日期    现住址
            //每一行数据，为表格中的一行
            contentRow = sheet.createRow(rowIndex);

            cell = contentRow.createCell(0);
            //为表格设置单元格的样式
            cell.setCellStyle(contentStyle);
            cell.setCellValue(user.getId());

            cell = contentRow.createCell(1);
            //为表格设置单元格的样式
            cell.setCellStyle(contentStyle);
            cell.setCellValue(user.getUserName());

            cell = contentRow.createCell(2);
            //为表格设置单元格的样式
            cell.setCellStyle(contentStyle);
            cell.setCellValue(user.getPhone());
            //处理日期的转换
            cell = contentRow.createCell(3);
            //为表格设置单元格的样式
            cell.setCellStyle(contentStyle);
            cell.setCellValue(dateFormat.format(user.getHireDate()));

            cell = contentRow.createCell(4);
            //为表格设置单元格的样式
            cell.setCellStyle(contentStyle);
            cell.setCellValue(user.getAddress());

            //行索引加一
            rowIndex++;
        }

        //导出文件：一个流（OutPutStream），两个头：文件打开方式（in-line,attachment）、文件下载的mime类型
        String fileName = "员工数据.xlsx";
        response.setHeader("Content-Disposition","attachment;filename="+new String(fileName.getBytes(),"ISO8859-1"));
        response.setContentType("pplication/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        //导出文件
        workbook.write(response.getOutputStream());
        workbook.close();

    }

    public void downLoadXlsxByPoiWithTemplate(HttpServletResponse response) throws Exception {
        //1.获取模板文件
        //获取到项目的根路径
        File rootPath = new File(ResourceUtils.getURL("classpath:").getPath());
        //获取根目录下的模板文件
        File templateFile = new File(rootPath, "/excel_template/userList.xlsx");
        //根据模板文件创建工作薄
        XSSFWorkbook workbook = new XSSFWorkbook(templateFile);

        //2.查询数据
        List<User> users = userMapper.selectAll();

        //获取工作薄中第二张工作表的样式，为当前的内容添加样式
        XSSFCellStyle contentStyle = workbook.getSheetAt(1).getRow(0).getCell(0).getCellStyle();

        //3.为模板添加数据
        //获取第一张工作表
        XSSFSheet sheet = workbook.getSheetAt(0);
        //从第三行开始添加数据
        Row contentRow = null;
        Cell cell = null;
        int rowIndex = 2;
        for (User user : users) {
            contentRow = sheet.createRow(rowIndex);
            //设置行高
            contentRow.setHeightInPoints(15.0F);

            cell = contentRow.createCell(0);
            //为单元格设置模板表格中的模板样式
            cell.setCellStyle(contentStyle);
            cell.setCellValue(user.getId());

            cell = contentRow.createCell(1);
            //为单元格设置模板表格中的模板样式
            cell.setCellStyle(contentStyle);
            cell.setCellValue(user.getUserName());

            cell = contentRow.createCell(2);
            //为单元格设置模板表格中的模板样式
            cell.setCellStyle(contentStyle);
            cell.setCellValue(user.getPhone());
            //处理日期的转换
            cell = contentRow.createCell(3);
            //为单元格设置模板表格中的模板样式
            cell.setCellStyle(contentStyle);
            cell.setCellValue(dateFormat.format(user.getHireDate()));

            cell = contentRow.createCell(4);
            //为单元格设置模板表格中的模板样式
            cell.setCellStyle(contentStyle);
            cell.setCellValue(user.getAddress());

            rowIndex++;
        }

        //删除模板中的第二张工作表，即导出是就只有一张表
        workbook.removeSheetAt(1);

        //4.导出Excel文件
        //导出文件：一个流（OutPutStream），两个头：文件打开方式（in-line,attachment）、文件下载的mime类型
        String fileName = "员工数据.xlsx";
        response.setHeader("Content-Disposition","attachment;filename="+new String(fileName.getBytes(),"ISO8859-1"));
        response.setContentType("pplication/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        //导出文件
        workbook.write(response.getOutputStream());
        workbook.close();

    }

    public void downLoadUserInfoByTemplate(Long id, HttpServletResponse response) throws Exception {
        //1.读取模板文件
        //获取项目的根路径
        File rootPath = new File(ResourceUtils.getURL("classpath:").getPath());
        File templateFile = new File(rootPath, "/excel_template/userinfo.xlsx");
        //根据模板创建工作薄
        XSSFWorkbook workbook = new XSSFWorkbook(templateFile);

        //2.获取单个用户的数据
        User user = userMapper.selectByPrimaryKey(id);
        //获取模板中的第一表
        XSSFSheet sheet = workbook.getSheetAt(0);

        //3.填充数据到模板
//        用户名：第2行，第2列
        sheet.getRow(1).getCell(1).setCellValue(user.getUserName());
//        手机号：第3行，第2列
        sheet.getRow(2).getCell(1).setCellValue(user.getPhone());
//        生日：第4行，第2列
        sheet.getRow(3).getCell(1).setCellValue(dateFormat.format(user.getBirthday()));
//        工资：第5行，第2列
        sheet.getRow(4).getCell(1).setCellValue(user.getSalary().toString());
//        入职日期：第6行，第2列
        sheet.getRow(5).getCell(1).setCellValue(dateFormat.format(user.getHireDate()));
//        司龄：第6行，第4列，司龄的公式处理：CONCATENATE(DATEDIF(B6,TODAY(),"Y"),"年",DATEDIF(B6,TODAY(),"YM"),"个月")
        //方法1，POI操作公式
//        sheet.getRow(5).getCell(3).setCellFormula("CONCATENATE(DATEDIF(B6,TODAY(),\"Y\"),\"年\",DATEDIF(B6,TODAY(),\"YM\"),\"个月\")");

        //方法2,在模板中定义公式后，直接调用，不用手动设置公式

//        省份：第7行，第2列
        sheet.getRow(6).getCell(1).setCellValue(user.getProvince());
//        城市：第7行，第4列
        sheet.getRow(6).getCell(3).setCellValue(user.getCity());
//        现住址：第8行，第2列
        sheet.getRow(7).getCell(1).setCellValue(user.getAddress());
//        照片：第2行到第5行，第3列到第4列
        //照片的处理
        //穿件字节输出流
        ByteArrayOutputStream byteArrayOutputStream = new ByteArrayOutputStream();
        //读取图片，将照片放入到了一个带缓存的类中
        BufferedImage bufferedImage = ImageIO.read(new File(rootPath,user.getPhoto()));

        //获取文件的扩展名
        String extName = user.getPhoto().substring(user.getPhoto().lastIndexOf(".") + 1).toUpperCase();
        System.out.println(extName);
        //将文件写出到字节输出流中，参数缓存图片类，文件的格式，流
        ImageIO.write(bufferedImage,extName,byteArrayOutputStream);

        //Patriarch类控制图片的写入，ClientAnchor，控制照片的位置
        XSSFDrawing patriarch = sheet.createDrawingPatriarch();
        //指定图片的位置：第2行到第5行，第3列到第4列
        //参数：0 左上角x轴的偏移量, 0 左上角y轴的偏移量, 0 右下角x轴的偏移量, 0 右下角y轴的偏移量, 0 起始列, 0 起始行, 0 结束列, 0 结束行
        //图片在结束列和结束行的左上角位置，要填充满的话要加 1，偏移单位1厘米=360000
        XSSFClientAnchor anchor = new XSSFClientAnchor(0, 0, 0, 0, 2, 1, 4, 5);

        int format = 0;
        switch (extName){
            case "JPG":
                format = XSSFWorkbook.PICTURE_TYPE_JPEG;
                break;
            case "JPEG":
                format = XSSFWorkbook.PICTURE_TYPE_JPEG;
                break;
            case "PNG":
                format = XSSFWorkbook.PICTURE_TYPE_PNG;
                break;
            case "GIF":
                format = XSSFWorkbook.PICTURE_TYPE_GIF;
                break;
            default:
                format = XSSFWorkbook.PICTURE_TYPE_JPEG;
        }
        //写出图片到工作表
        patriarch.createPicture(anchor,workbook.addPicture(byteArrayOutputStream.toByteArray(),format));

        //4.导出Excel文件
        //导出文件：一个流（OutPutStream），两个头：文件打开方式（in-line,attachment）、文件下载的mime类型
        String fileName = "员工("+user.getUserName()+")详细数据.xlsx";
        response.setHeader("Content-Disposition","attachment;filename="+new String(fileName.getBytes(),"ISO8859-1"));
        response.setContentType("pplication/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        //导出文件
        workbook.write(response.getOutputStream());
        workbook.close();
    }

    public void downLoadUserInfoByTemplateEngine(Long id, HttpServletResponse response) throws Exception {
        //1.读取模板文件
        //获取项目的根路径
        File rootPath = new File(ResourceUtils.getURL("classpath:").getPath());
        File templateFile = new File(rootPath, "/excel_template/userinfoTemplateEngine.xlsx");
        //根据模板创建工作薄
        org.apache.poi.ss.usermodel.Workbook workbook = new XSSFWorkbook(templateFile);

        //2.获取单个用户的数据
        User user = userMapper.selectByPrimaryKey(id);
        //使用工具类来给模板引擎添加数据：参数：单个用户数据，workbook，图片的绝对路径
        workbook = ExcelExportEngine.writeToExcel(user, workbook,rootPath+user.getPhoto());

        //4.导出Excel文件
        //导出文件：一个流（OutPutStream），两个头：文件打开方式（in-line,attachment）、文件下载的mime类型
        String fileName = "员工("+user.getUserName()+")详细数据.xlsx";
        response.setHeader("Content-Disposition","attachment;filename="+new String(fileName.getBytes(),"ISO8859-1"));
        response.setContentType("pplication/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        //导出文件
        workbook.write(response.getOutputStream());
        workbook.close();

    }

    /**
     * 百万数据导出：1.肯定使用高版本的Excel、2.使用SAX解析Excel(xml)
     * 限制：1.不能使用模板，2.不能使用太多的样式
     */
    public void downLoadMillion(HttpServletResponse response) throws Exception {
        //1.创建一个使用SAX解析Excel的工作薄，下面使用Sax解析Excel(xml)，是逐行执行，时间较长不会出现内存溢出异常
        SXSSFWorkbook workbook = new SXSSFWorkbook();
        //下面使用dom4j，方式解析，直接将Excel全部数据加载，会使堆内存溢出报异常信息
//        XSSFWorkbook workbook = new XSSFWorkbook();
        //导出500W数据，不可能放到一张sheet中，要放到5张sheet中，每张100W,在高版本EXCEl中一个sheet最多可以存100多W条数据

        //2.循环查询百万数据
        int page = 1;
        //表示当前处理数据的条数
        int num = 0;
        //记录每个sheet的数据行索引,每次在第二行插入数据
        int rowIndex = 1;
        //节省栈内存
        Row row = null;
        //提取sheet,导出数据时要操作，放到while循环中会出现null的情况，要放到外面
        SXSSFSheet sheet = null;
        while(true){
            //每张工作表方10W条数据
            List<User> userList = this.findPage(page, 100000);
            if (CollectionUtils.isEmpty(userList)){
                //用户数据为空退出循环
                break;
            }
            // 0 1000000 2000000 3000000 4000000 5000000
            //编号	姓名	手机号	入职日期	现住址
            //导出5张表，没张表100W条数据
            if(num % 1000000 == 0){
                //创建工作表，后rowIndex，sheet行索引值重置为1
                sheet = workbook.createSheet("第"+((num/100000)+1)+"个工作表");
                rowIndex = 1;
                //新的工作表要创建标题
                String[] titles = {"编号","姓名","手机号","入职日期","现住址"};
                //标题固定在第一行
                SXSSFRow titleRow = sheet.createRow(0);
                for (int i = 0; i < 5; i++) {
                    //第一行创建5个单元格，并设置标题的值
                   titleRow.createCell(i).setCellValue(titles[i]);
                }
            }

            for (User user : userList) {
                //创建行
                row = sheet.createRow(rowIndex);

                row.createCell(0).setCellValue(user.getId());
                row.createCell(1).setCellValue(user.getUserName());
                row.createCell(2).setCellValue(user.getPhone());
                row.createCell(3).setCellValue(dateFormat.format(user.getHireDate()));
                row.createCell(4).setCellValue(user.getAddress());

                //每条数据处理完，行加1,数据条数加1，记录当前sheet的数量
                rowIndex++;
                //记录当前工作薄的所有记录数
                num++;
            }
            //页码加1
            page++;
        }

        //4.导出Excel文件
        //导出文件：一个流（OutPutStream），两个头：文件打开方式（in-line,attachment）、文件下载的mime类型
        String fileName = "百万用户数据的导出.xlsx";
        response.setHeader("Content-Disposition","attachment;filename="+new String(fileName.getBytes(),"ISO8859-1"));
        response.setContentType("pplication/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        //导出文件
        workbook.write(response.getOutputStream());
        workbook.close();

    }

    public void downLoadCSV(HttpServletResponse response) throws Exception {

        ServletOutputStream outputStream = response.getOutputStream();
        //导出文件：一个流（OutPutStream），两个头：文件打开方式（in-line,attachment）、文件下载的mime类型
        String fileName = "百万用户数据的导出.csv";
        response.setHeader("Content-Disposition","attachment;filename="+new String(fileName.getBytes(),"ISO8859-1"));
        response.setContentType("text/csv");
        //使用openCSV,创建CsvWriter类，来导出文件
        CSVWriter csvWriter = new CSVWriter(new OutputStreamWriter(outputStream, "utf-8"));

        //写入固定的标题
        String[] title = {"编号","姓名","手机号","入职日期","现住址"};
        csvWriter.writeNext(title);
        int page =1;
        while (true){
            //查找数据库中的数据，每次查找20W条数据,为空则跳出循环
            List<User> userList = this.findPage(page, 200000);
            if (CollectionUtils.isEmpty(userList)){
                break;
            }
            for (User user : userList) {
                //循环查询的数据后，封装为一个数组，使用csvWriter一条一条导出，
                csvWriter.writeNext(new String[]{user.getId().toString(),user.getUserName(),user.getPhone(),
                        dateFormat.format(user.getHireDate()),user.getAddress()});
            }
            page++;
            //清空缓存
            csvWriter.flush();
        }
        csvWriter.close();
        outputStream.close();
    }


    //下载用户合同
    public void downloadContract(Long id, HttpServletResponse response) throws Exception {
        //1.读取模板
        File rootPath = new File(ResourceUtils.getURL("classpath:").getPath());
        File templateFile = new File(rootPath, "/word_template/contract_template.docx");
        //创建XWPFDocument，读取模板文件
        XWPFDocument word = new XWPFDocument(new FileInputStream(templateFile));

        //2.查询用户数据和resource的数据，不使用userMapper查询数据，将bean转换为map，后面好方便比较
        User user = this.findbyId(id);
        Map<String,String> map = new HashMap<>();
        map.put("userName",user.getUserName());
        map.put("hireDate", dateFormat.format(user.getHireDate()));
        map.put("address",user.getAddress());

        //3.向模板替换数据
        //处理正文
        //先获取模板中的所有段落，片段，通过最小单元和map的键比较如果相等，就替换值
        List<XWPFParagraph> paragraphs = word.getParagraphs();
        for (XWPFParagraph paragraph : paragraphs) {
            List<XWPFRun> runs = paragraph.getRuns();
            for (XWPFRun run : runs) {
                String text = run.getText(0);
                for (String key : map.keySet()){
                    if (key.equals(text)){
                        run.setText(text.replaceAll(key,map.get(key)),0);
                    }
                }
            }
        }

        //处理表格  名称  价值  价格  是否需要归还  照片
        //获取用户的Resource来创建表格中的数据
        List<Resource> resourceList = user.getResourceList();
        //获取文档中的第一张表格
        XWPFTable table = word.getTables().get(0);

        //获取表格中的第一行
        XWPFTableRow row = table.getRow(0);
        int rowIndex = 1;
        for (Resource resource : resourceList) {
            //添加行，不使用insertNewTableRow()方法如果用来则插入的表格行没有任何的样式，要自己设置，要复杂，使用add直接复制样式简单
            //addRow()是浅复制，操作的还是原来的对象即还是原来的行，数据会有所错乱，手写方法进行深克隆行
//            table.addRow(row);
            copyRow(table,row,rowIndex);
            XWPFTableRow row1 = table.getRow(rowIndex);
            row1.getCell(0).setText(resource.getName());
            row1.getCell(1).setText(resource.getPrice().toString());
            row1.getCell(2).setText(resource.getNeedReturn()?"需要":"不需要");
            //处理导出的图片,获取图片的路径
            File imageFile = new File(rootPath,"/static"+resource.getPhoto());
            setCellImage(row1.getCell(3),imageFile);

            //行索引加1
            rowIndex++;
        }

        //4.导出word文件
        //导出文件：一个流（OutPutStream），两个头：文件打开方式（in-line,attachment）、文件下载的mime类型
        String fileName = "员工("+user.getUserName()+")合同.docx";
        response.setHeader("Content-Disposition","attachment;filename="+new String(fileName.getBytes(),"ISO8859-1"));
        response.setContentType("pplication/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        //导出文件
        word.write(response.getOutputStream());

    }

    //导出word合同，为单元格设置图片
    private void setCellImage(XWPFTableCell cell, File imageFile) {
        //创建word的最小单元
        XWPFRun run = cell.getParagraphs().get(0).createRun();

        //将inputStream写到try里面，程序可以自动关闭流，不需要手动关闭流
        try(FileInputStream inputStream = new FileInputStream(imageFile);) {
//        InputStream pictureData 流, int pictureType图片的类型, String filename 文件名, int width 宽, int height 高
            //使用单位时，要调用方法，底层是将输入的值进行扩大，否则的话直接设置值太小，导出后图片不显示
            run.addPicture(inputStream,XWPFDocument.PICTURE_TYPE_JPEG,imageFile.getName(), Units.toEMU(100),Units.toEMU(50));
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    /**导出word合同，复制表格行的样式
     * 深克隆表格行的样式
     * @param table 表格
     * @param sourceRow 要复制的行的样式
     * @param rowIndex 行的索引
     */
    private void copyRow(XWPFTable table, XWPFTableRow sourceRow, int rowIndex) {
        //插入一个新的行到指定的索引处，完全没有任何的样式
        XWPFTableRow targetRow = table.insertNewTableRow(rowIndex);
        //为该行设置和复制，他之前的行的样式和属性(tr行，pr属性)
        targetRow.getCtRow().setTrPr(sourceRow.getCtRow().getTrPr());

        //获取源行的单元格
        List<XWPFTableCell> cells = sourceRow.getTableCells();
        //为空则处理
        if(CollectionUtils.isEmpty(cells)){
            return;
        }

        //节省栈内存
        XWPFTableCell targetCell = null;
        for (XWPFTableCell cell : cells) {
            //添加单元格
            targetCell = targetRow.addNewTableCell();
            //复制单元格的样式
            targetCell.getCTTc().setTcPr(cell.getCTTc().getTcPr());
            //复制单元格中段落的样式
            targetCell.getParagraphs().get(0).getCTP().setPPr(cell.getParagraphs().get(0).getCTP().getPPr());
        }

    }

    //使用EasyPoI导出Excel
    public void downLoadWithEasyPOI(HttpServletResponse response) throws Exception {
        //设置导出的Excel文件的类型和文件的格式（具体看源码）
        ExportParams exportParams = new ExportParams("员工信息列表","sheet1", ExcelType.XSSF);

        //查询用户的数据
        List<User> userList = userMapper.selectAll();

        //使用EasyPOI导出
        org.apache.poi.ss.usermodel.Workbook workbook = ExcelExportUtil.exportExcel(exportParams, User.class, userList);

        //导出文件：一个流（OutPutStream），两个头：文件打开方式（in-line,attachment）、文件下载的mime类型
        String fileName = "用户数据的导出.xlsx";
        response.setHeader("Content-Disposition","attachment;filename="+new String(fileName.getBytes(),"ISO8859-1"));
        response.setContentType("pplication/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        //导出文件
        workbook.write(response.getOutputStream());
        workbook.close();

    }

    //使用EasyPOI导入Excel
    public void uploadExcelWithEasyPOI(MultipartFile file) throws Exception {
        //设置导入的文件的参数和文件的格式
        ImportParams importParams = new ImportParams();
        //设置文件上传后不保存
        importParams.setNeedSave(false);
        //设置文件的大标题和小标额占一行
        importParams.setTitleRows(1);
        importParams.setHeadRows(1);
        List<User> userList = ExcelImportUtil.importExcel(file.getInputStream(), User.class, importParams);

        for (User user : userList) {
            //消除mysql中Id主键的异常
            user.setId(null);
            userMapper.insert(user);
        }

    }

    public void downLoadUserInfoByTemplateEngineWithEasyPOI(Long id, HttpServletResponse response) throws Exception {
        //1.读取模板文件
        //获取项目的根路径
        File rootPath = new File(ResourceUtils.getURL("classpath:").getPath());
        File templateFile = new File(rootPath, "/excel_template/userinfoTemplateEngineToEasy.xlsx");

        //获取模板引擎后，设置导出模板的参数
        TemplateExportParams exportParams = new TemplateExportParams(templateFile.getPath(),true);
        //查找用户数据
        User user = userMapper.selectByPrimaryKey(id);
        //将Bean转map
        Map<String, Object> map = EntityUtils.entityToMap(user);
        //处理图片
        ImageEntity imageEntity = new ImageEntity();
        imageEntity.setUrl(user.getPhoto());
        //设置图片占几列
        imageEntity.setColspan(2);
        //设置图片占几行
        imageEntity.setRowspan(4);
        //加入到map
        map.put("photo",imageEntity);

        org.apache.poi.ss.usermodel.Workbook workbook = ExcelExportUtil.exportExcel(exportParams, map);

        //导出文件：一个流（OutPutStream），两个头：文件打开方式（in-line,attachment）、文件下载的mime类型
        String fileName = "用户("+user.getUserName()+")数据.xlsx";
        response.setHeader("Content-Disposition","attachment;filename="+new String(fileName.getBytes(),"ISO8859-1"));
        response.setContentType("pplication/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        //导出文件
        workbook.write(response.getOutputStream());
        workbook.close();

    }

    public void downLoadCSVWithEasyPOI(HttpServletResponse response) throws Exception {

        ServletOutputStream outputStream = response.getOutputStream();
        //导出文件：一个流（OutPutStream），两个头：文件打开方式（in-line,attachment）、文件下载的mime类型
        String fileName = "用户数据的导出.csv";
        response.setHeader("Content-Disposition","attachment;filename="+new String(fileName.getBytes(),"ISO8859-1"));
        response.setContentType("text/csv");

        //设置导出参数Easy
        CsvExportParams exportParams = new CsvExportParams();
        //导出数据的时候排除照片的导出，CSV为文本的格式，排除即Bean注解上面的name
        exportParams.setExclusions(new String[]{"照片"});

        //查用户数据
        List<User> userList = userMapper.selectAll();
        CsvExportUtil.exportCsv(exportParams,User.class,userList,outputStream);
    }

    public void downloadContractWithEasyPOI(Long id, HttpServletResponse response) throws Exception {
        //1.读取模板文件
        //获取项目的根路径
        File rootPath = new File(ResourceUtils.getURL("classpath:").getPath());
        File templateFile = new File(rootPath, "/word_template/contract_templateToEasy.docx");

        //查找数据
        User user = this.findbyId(id);
        //将用户数据封装成一个map
        Map<String,Object> params = new HashMap<>();
        params.put("userName", user.getUserName());
        params.put("hireDate", dateFormat.format(user.getHireDate()));
        params.put("address", user.getAddress());

        //默认使用EasyPOI导出word文档的表格中，是不可以插入图片的，值可以在段落出插入图片，一下就是验证
        //所有EasyPOI并不是万能的，一下复杂的功能还是要使用POI来操作
        ImageEntity imageEntityContent = new ImageEntity();
        imageEntityContent.setUrl(rootPath.getPath()+user.getPhoto());
        imageEntityContent.setWidth(100);
        imageEntityContent.setHeight(50);
        params.put("photo",imageEntityContent);


        //将用户的resource封装成map,后放到一个集合中
        List<Map>  resourceList= new ArrayList<>() ;
        Map<String,Object> map = null;
        for (Resource resource : user.getResourceList()) {
            map = new HashMap<>();
            map.put("name",resource.getName());
            map.put("price",resource.getName());
            map.put("needReturn",resource.getName());
            //处理照片
            ImageEntity imageEntity = new ImageEntity();
            imageEntity.setUrl(rootPath+"/static/"+resource.getPhoto());
            map.put("photo",imageEntity);

            resourceList.add(map);
        }
        //将list集合放到参数集合中
        params.put("resourceList",resourceList);

        //填充模板数据
        XWPFDocument word = WordExportUtil.exportWord07(templateFile.getPath(), params);

        //导出word文件
        //导出文件：一个流（OutPutStream），两个头：文件打开方式（in-line,attachment）、文件下载的mime类型
        String fileName = "员工("+user.getUserName()+")合同.docx";
        response.setHeader("Content-Disposition","attachment;filename="+new String(fileName.getBytes(),"ISO8859-1"));
        response.setContentType("pplication/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        //导出文件
        word.write(response.getOutputStream());
    }

    public void downLoadPDF(HttpServletResponse response) throws Exception {
        //获取模板文件
        File rootPath = new File(ResourceUtils.getURL("classpath:").getPath());
        File templateFile = new File(rootPath, "/pdf_template/userList-db.jasper");

        //准备数据库的连接
        Map params = new HashMap();
        //读取模板文件，因为模板文件用的是Filed没有Parameters，但要传一个map，最后为一个数据库的连接，可以手写，推荐使用数据源来获取连接
        JasperPrint jasperPrint = JasperFillManager.fillReport(new FileInputStream(templateFile), params, hikariDataSource.getConnection());

        //导出PDF文件
        ServletOutputStream outputStream = response.getOutputStream();
        //导出文件：一个流（OutPutStream），两个头：文件打开方式（in-line,attachment）、文件下载的mime类型
        String fileName = "用户列表数据.pdf";
        response.setHeader("Content-Disposition","attachment;filename="+new String(fileName.getBytes(),"ISO8859-1"));
        response.setContentType("application/pdf");

        //导出
        JasperExportManager.exportReportToPdfStream(jasperPrint,outputStream);
    }

    public void downLoadPDFHandleDate(HttpServletResponse response) throws Exception {
        //获取模板文件
        File rootPath = new File(ResourceUtils.getURL("classpath:").getPath());
        //使用没用按省份分组的模板导出
//        File templateFile = new File(rootPath, "/pdf_template/userList.jasper");

        //使用按省份模板的pdf导出
        File templateFile = new File(rootPath, "/pdf_template/userListGroup.jasper");

        //准备数据库的连接
        Map params = new HashMap();
        //设置按省份来排序，否则会出现省份重复的问题
        Example example = new Example(User.class);
        example.setOrderByClause("province");
        List<User> userList = userMapper.selectByExample(example);

        //使用流式编程和lambda表达式给hireDateStr赋值后，在重新给userList赋值
        userList = userList.stream().map(user -> {
            user.setHireDateStr(dateFormat.format(user.getHireDate()));
            return user;
        }).collect(Collectors.toList());

        //将后台查询的数据转换为JasperDatasource
        JRBeanCollectionDataSource dataSource = new JRBeanCollectionDataSource(userList);

        //读取模板文件，因为模板文件用的是Filed没有Parameters，但要传一个map，最后为一个数据库的连接，可以手写，推荐使用数据源来获取连接
        JasperPrint jasperPrint = JasperFillManager.fillReport(new FileInputStream(templateFile), params,dataSource);

        //导出PDF文件
        ServletOutputStream outputStream = response.getOutputStream();
        //导出文件：一个流（OutPutStream），两个头：文件打开方式（in-line,attachment）、文件下载的mime类型
        String fileName = "用户列表数据.pdf";
        response.setHeader("Content-Disposition","attachment;filename="+new String(fileName.getBytes(),"ISO8859-1"));
        response.setContentType("application/pdf");

        //导出
        JasperExportManager.exportReportToPdfStream(jasperPrint,outputStream);
    }

    public void downLoadUserInfoWithJasperPDF(Long id, HttpServletResponse response) throws Exception {
        //获取模板文件
        File rootPath = new File(ResourceUtils.getURL("classpath:").getPath());
        File templateFile = new File(rootPath, "/pdf_template/userInfo.jasper");

        //查询单个用户的信息后，封装为map，因为模板中使用的是Parameter来填充数据，使用Field来填充集合数据
        User user = userMapper.selectByPrimaryKey(id);
        //转map时同时处理了日期格式的问题
        Map params = EntityUtils.entityToMap(user);
        //处理工资格式问题，模板中为字符串，map中为Integer
        params.put("salary",user.getSalary().toString());
        //处理图片的路径问题
        params.put("photo",rootPath+user.getPhoto());

        //读取模板文件，因为模板文件用的是Filed没有Parameters，但要传一个map，最后为一个数据库的连接，可以手写，推荐使用数据源来获取连接
        JasperPrint jasperPrint = JasperFillManager.fillReport(new FileInputStream(templateFile), params,new JREmptyDataSource());

        //导出PDF文件
        ServletOutputStream outputStream = response.getOutputStream();
        //导出文件：一个流（OutPutStream），两个头：文件打开方式（in-line,attachment）、文件下载的mime类型
        String fileName = "用户详细数据.pdf";
        response.setHeader("Content-Disposition","attachment;filename="+new String(fileName.getBytes(),"ISO8859-1"));
        response.setContentType("application/pdf");

        //导出
        JasperExportManager.exportReportToPdfStream(jasperPrint,outputStream);

    }
}

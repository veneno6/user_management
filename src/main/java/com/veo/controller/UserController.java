package com.veo.controller;

import com.veo.pojo.User;
import com.veo.service.UserService;
import org.jfree.chart.ChartFactory;
import org.jfree.chart.ChartUtils;
import org.jfree.chart.JFreeChart;
import org.jfree.chart.StandardChartTheme;
import org.jfree.data.category.DefaultCategoryDataset;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;
import java.awt.*;
import java.io.File;
import java.util.List;

@RestController
@RequestMapping("/user")
public class UserController {

    @Autowired
    private UserService userService;

    @GetMapping("/findPage")
    public List<User>  findPage(
            @RequestParam(value = "page",defaultValue = "1") Integer page,
            @RequestParam(value = "rows",defaultValue = "10") Integer pageSize){
        return userService.findPage(page,pageSize);
    }

    @GetMapping(value = "/downLoadXlsByJxl",name = "使用jxl导出Excel")
    public void downLoadXlsByJxl(HttpServletResponse response) throws Exception {
        userService.downLoadXlsByJxl(response);
    }

    @PostMapping(value = "/uploadExcel",name = "使用POI导入Excel")
    public void uploadExcel(MultipartFile file) throws Exception {
        //使用POI导入Excel
//        userService.uploadExcel(file);

        //使用EasyPOI导入Excel
        userService.uploadExcelWithEasyPOI(file);
    }

    @GetMapping(value = "/downLoadXlsxByPoi",name = "使用POI导出所有用户数据Excel")
    public void downLoadXlsxByPoi(HttpServletResponse response) throws Exception {
        //基本POI导出Excel，不带任何的样式
//        userService.downLoadXlsxByPoi(response);
        //带样式的POI导出
//        userService.downLoadXlsxByPoiWithCellStyle(response);
        //使用带模板的Excel导出
        userService.downLoadXlsxByPoiWithTemplate(response);
    }

    @GetMapping(value = "/download",name = "使用POI导出用户详细信息Excel")
    public void downLoadUserInfoByTemplate(Long id,HttpServletResponse response) throws Exception {
        //使用模板来导出单个用户数据，模板数据死板
//        userService.downLoadUserInfoByTemplate(id,response);

        //使用自定义的模板引擎来导出用户数据，模板数据可动，Bug每次启动需要先删除targe文件，否则导出的是上次的数据，（值传递问题）
//        userService.downLoadUserInfoByTemplateEngine(id,response);

        //使用EasyPOI导出用户详细信息
//        userService.downLoadUserInfoByTemplateEngineWithEasyPOI(id,response);

        //使用Jasper导出用户的详细数据为PDF格式
        userService.downLoadUserInfoWithJasperPDF(id,response);
    }

    @GetMapping(value = "/downLoadMillion",name = "使用POI导出百万用户数据")
    public void downLoadMillion(HttpServletResponse response) throws Exception {
        userService.downLoadMillion(response);
    }

    @GetMapping(value = "/downLoadCSV",name = "使用CSV导出百万用户数据")
    public void downLoadCSV(HttpServletResponse response) throws Exception {
        //使用CSV导出百万用户数据
//        userService.downLoadCSV(response);
        //使用EasyPOI导出用户数据CSV格式，Easy导出不了百万级别的数据
        userService.downLoadCSVWithEasyPOI(response);
    }

    @GetMapping(value = "/{id}",name = "根据id查用户数据")
    public User findById(@PathVariable("id") Long id) throws Exception {
        return userService.findbyId(id);
    }

    @GetMapping(value = "/downloadContract",name = "下载用户合同")
    public void downloadContract(Long id,HttpServletResponse response) throws Exception {
        //使用POI导出Word
//        userService.downloadContract(id,response);
        //使用EasyPOI导出word
        userService.downloadContractWithEasyPOI(id,response);
    }

    @GetMapping(value = "/downLoadWithEasyPOI",name = "使用easyPOI导出Excel")
    public void downLoadWithEasyPOI(HttpServletResponse response) throws Exception {
        userService.downLoadWithEasyPOI(response);
    }

    @GetMapping(value = "/downLoadPDF",name = "导出用户数据到PDF中")
    public void downLoadPDF(HttpServletResponse response) throws Exception {
        //导出用户列表数据没有出路日期的格式
//        userService.downLoadPDF(response);
        //导出用户数据，处理了日期的格式
        userService.downLoadPDFHandleDate(response);

    }

    @GetMapping(value = "/jfreeChart",name = "显示一张JfreeChart图片")
    public void jfreeChart(HttpServletResponse response) throws Exception {
        //生成柱状图的数据
        //每年个部门的入职人数
        DefaultCategoryDataset dataset = new DefaultCategoryDataset();
        dataset.addValue(13,"人事部","2011");
        dataset.addValue(64,"人事部","2012");
        dataset.addValue(34,"人事部","2013");
        dataset.addValue(64,"人事部","2014");
        dataset.addValue(34,"人事部","2015");

        dataset.addValue(43,"技术部","2011");
        dataset.addValue(56,"技术部","2012");
        dataset.addValue(24,"技术部","2013");
        dataset.addValue(36,"技术部","2014");
        dataset.addValue(63,"技术部","2015");

        dataset.addValue(23,"销售部","2011");
        dataset.addValue(34,"销售部","2012");
        dataset.addValue(55,"销售部","2013");
        dataset.addValue(34,"销售部","2014");
        dataset.addValue(44,"销售部","2015");

        dataset.addValue(43,"财务部","2011");
        dataset.addValue(23,"财务部","2012");
        dataset.addValue(45,"财务部","2013");
        dataset.addValue(26,"财务部","2014");
        dataset.addValue(43,"财务部","2015");

        dataset.addValue(45,"法务部","2011");
        dataset.addValue(35,"法务部","2012");
        dataset.addValue(32,"法务部","2013");
        dataset.addValue(25,"法务部","2014");
        dataset.addValue(35,"法务部","2015");

        //设置图形的主题，解决中文不显示问题
        StandardChartTheme chartTheme = new StandardChartTheme("CN");

        //设置大标题的字体样式
        chartTheme.setExtraLargeFont(new Font("华文宋体",Font.BOLD,20));
        //设置图例的字体样式
        chartTheme.setRegularFont(new Font("华文宋体",Font.BOLD,15));
//        设置内容的字体样式x,y轴
        chartTheme.setLargeFont(new Font("华文宋体",Font.BOLD,15));

        //设置样式
        ChartFactory.setChartTheme(chartTheme);

        //参数：标题 x轴的描述 y轴的描述 数据
        JFreeChart pieChart = ChartFactory.createBarChart("公司人数","各部门","入职人数",dataset);

        ChartUtils.writeChartAsJPEG(response.getOutputStream(),pieChart,400,300);

    }

}

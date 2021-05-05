package com.charts;

import org.jfree.chart.ChartFactory;
import org.jfree.chart.ChartUtils;
import org.jfree.chart.JFreeChart;
import org.jfree.chart.StandardChartTheme;
import org.jfree.data.category.DefaultCategoryDataset;
import org.jfree.data.general.DefaultPieDataset;
import org.junit.Test;

import java.awt.*;
import java.io.File;
import java.io.IOException;

public class JfreeChatsDemo {

    @Test
    public void testPieChart() throws IOException {
        //生成饼状图的数据
        DefaultPieDataset dataset = new DefaultPieDataset();
        dataset.setValue("销售部",300);
        dataset.setValue("技术部",500);
        dataset.setValue("人事部",800);

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

        //String title 图的标题, PieDataset dataset 要显示的数据, boolean legend 图例, boolean tooltips 鼠标悬浮显示提示, boolean urls 点击跳转
        //2D饼状图
//        JFreeChart pieChart = ChartFactory.createPieChart("各部门人数", dataset, true, false, false);
        //3D饼状图
        JFreeChart pieChart = ChartFactory.createPieChart3D("各部门人数", dataset, true, false, false);

        ChartUtils.saveChartAsPNG(new File("d:\\pieChart.png"),pieChart,400,300);
    }


    @Test
    public void testLineChart() throws IOException {
        //生成线性图的数据
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
        JFreeChart pieChart = ChartFactory.createLineChart("公司人数","各部门","入职人数",dataset);

        ChartUtils.saveChartAsPNG(new File("d:\\lineChart.png"),pieChart,400,300);
    }

    @Test
    public void testBarChart() throws IOException {
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

        ChartUtils.saveChartAsPNG(new File("d:\\barChart.png"),pieChart,400,300);
    }


}

package com.veo;

import net.sf.jasperreports.engine.JREmptyDataSource;
import net.sf.jasperreports.engine.JasperExportManager;
import net.sf.jasperreports.engine.JasperFillManager;
import net.sf.jasperreports.engine.JasperPrint;
import org.junit.Test;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.HashMap;
import java.util.Map;

//使用Jaspersoft Studio 编译好的模板文件导出PDF
public class JasperPDFDemo {
    //测试英文
    public static void main(String[] args) throws Exception {
        //获取pdf文件的模板，使用Jaspersoft Studio 软件制作模板
        FileInputStream inputStream = new FileInputStream("test01-en.jasper");

        //构造模板使用的参数文件，默认所有中文都不显示，要配置具体的字体
        Map params = new HashMap<>();
        params.put("userNameP","zhangsan");
        params.put("phoneP","13800000000");

        //向模板文件中填充数据信息，参数：模板流文件，数据信息，数据源
        JasperPrint jasperPrint = JasperFillManager.fillReport(inputStream, params, new JREmptyDataSource());

        //导出pdf文件
        JasperExportManager.exportReportToPdfStream(jasperPrint,new FileOutputStream("test01-en.pdf"));
    }

    //测试中文,其中模板文件中的字体的字形要改为华文宋体，并导入字体资源到resource目录下后编译生成target文件
    @Test
    public void testJasperCh() throws Exception {
        //获取pdf文件的模板，使用Jaspersoft Studio 软件制作模板
        FileInputStream inputStream = new FileInputStream("test01-ch.jasper");

        //构造模板使用的参数文件，默认所有中文都不显示，要配置具体的字体
        //使用的Parameter参数，（Field数据库中的字段）
        Map params = new HashMap<>();
        params.put("userNameP","张三");
        params.put("phoneP","13800000000");

        //向模板文件中填充数据信息，参数：模板流文件，数据信息，数据源
        JasperPrint jasperPrint = JasperFillManager.fillReport(inputStream, params, new JREmptyDataSource());

        //导出pdf文件
        JasperExportManager.exportReportToPdfStream(jasperPrint,new FileOutputStream("test01-ch.pdf"));
    }
}

package com.veo;

import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.Dispatch;

//只用Jacob调用本地主机的Office软件，来讲Word文件打开后另存为Pdf文件，前提是本地主句安装了Office软件
//如果出现VariantChangeType failed错误要修改word DCOM默认的标识，改为“交互式用户”模式，即可正常调用了。（百度）
//修改如果默认报错，但是文件转换成功即可
public class JacobDeamoToPDF {
    public static void main(String[] args) {
        //具体路径需要修改
        String source = "C:\\Users\\veneno\\Desktop\\study\\report_froms\\user_management\\张三_合同.docx";
        String target = "C:\\Users\\veneno\\Desktop\\study\\report_froms\\user_management\\张三_合同.pdf";
        ActiveXComponent app = null;
        try {
            //使用jacob中的ActiveXComponent，来操作Office中的软件，下面操作Word软件
            app = new ActiveXComponent("Word.application");

            //调用Word时不显示窗口
            app.setProperty("Visible",false);
            //获得打开的所有Word文档
            Dispatch docs = app.getProperty("Documents").toDispatch();
            //使用Open命令从所有文档中打开指定的文档
            Dispatch doc = Dispatch.call(docs, "Open", source).toDispatch();

            // 另存为，将文档保存为pdf，其中Word保存为pdf的格式宏的值是17
            Dispatch.call(doc, "SaveAS", target, 17).toDispatch();
            //关闭打开的文件
            Dispatch.call(doc,"Close");
        } catch (Exception e) {
            System.out.println("Word转PDF出错：" + e.getMessage());
        } finally {
            // 关闭office
            if (app != null) {
                app.invoke("Quit", 0);
            }
        }


    }
}

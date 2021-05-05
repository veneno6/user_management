package com.veo.pojo;
import cn.afterturn.easypoi.excel.annotation.Excel;
import com.fasterxml.jackson.annotation.JsonFormat;
import com.fasterxml.jackson.annotation.JsonIgnore;
import lombok.Data;
import tk.mybatis.mapper.annotation.KeySql;

import javax.persistence.Id;
import javax.persistence.Table;
import javax.persistence.Transient;
import java.util.Date;
import java.util.List;

/**
 * 员工
 */
@Data
//导出普通的数据表
@Table(name="tb_user")
//百万数据导出数据库中的表（csv同样）
//@Table(name="tb_user2")
public class User {
    @Id
    @KeySql(useGeneratedKeys = true)
    @Excel(name = "编号",orderNum = "0",width = 15)
    private Long id;         //主键
    @Excel(name = "姓名",orderNum = "1",width = 15,isImportField = "true")
    private String userName; //员工名
    @Excel(name = "手机号",orderNum = "2",width = 15,isImportField = "true")
    private String phone;    //手机号
    @Excel(name = "省份",orderNum = "3",width = 15,isImportField = "true")
    private String province; //省份名
    @Excel(name = "城市",orderNum = "4",width = 15,isImportField = "true")
    private String city;     //城市名
    @Excel(name = "工资",orderNum = "5",width = 10,type = 10,isImportField = "true")
    private Integer salary;   // 工资
    @Excel(name = "入职日期",orderNum = "6",width = 15,format = "yyyy-MM-dd",isImportField = "true")
    @JsonFormat(pattern="yyyy-MM-dd")
    private Date hireDate; // 入职日期

    //处理到处pdf时的日期问题，不与数据库管联
    @JsonIgnore //转json不考虑这个字段
    @Transient
    private String hireDateStr;

    private String deptId;   //部门id
    @Excel(name = "出生日期",orderNum = "7",width = 15,format = "yyyy-MM-dd",isImportField = "true")
    private Date birthday; //出生日期
    @Excel(name = "照片",orderNum = "9",width = 15,type = 2,isImportField = "true",savePath = "C:\\Users\\veneno\\Desktop\\study\\report_froms\\user_management\\src\\main\\resources\\static\\user_photos")
    private String photo;    //一寸照片
    @Excel(name = "现住址",orderNum = "8",width = 28,isImportField = "true")
    private String address;  //现在居住地址

    private List<Resource> resourceList; //办公用品

    @Override
    public String toString() {
        return "User{" +
                "id=" + id +
                ", userName='" + userName + '\'' +
                '}';
    }
}

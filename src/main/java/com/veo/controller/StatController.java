package com.veo.controller;

import com.veo.service.StatService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import java.util.List;
import java.util.Map;

@RestController
@RequestMapping("/stat")
public class StatController {

    @Autowired
    private StatService statService;

//    res.data = [{deptName:"财务部",num:11},{},{}]
    @RequestMapping(value = "/columnCharts",name = "统计各部门人数")
    public List<Map> columnCharts(){
        return statService.columnCharts();
    }

    //res.data = [{name:"01",num:2},{},{}]
    @RequestMapping(value = "/lineCharts",name = "月份入职人数统计")
    public List<Map> lineCharts(){
        return statService.lineCharts();
    }

    // 员工地方来源统计 pieCharts()
    @RequestMapping(value = "/pieCharts",name = "员工地方来源统计")
    public List<Map<String,Object>> pieCharts(){
        return statService.pieCharts();
    }

    // 员工地方来源统计 pieCharts()
    @RequestMapping(value = "/pieECharts",name = "员工地方来源统计Echarts方式")
    public Map pieECharts(){
        return statService.pieECharts();
    }


}

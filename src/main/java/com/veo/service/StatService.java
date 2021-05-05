package com.veo.service;

import com.veo.mapper.UserMapper;
import com.veo.pojo.User;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import tk.mybatis.mapper.entity.Example;

import java.util.*;
import java.util.stream.Collectors;

@Service
public class StatService {

    @Autowired
    private UserMapper userMapper;


    public List<Map> columnCharts() {
        return userMapper.columnCharts();
    }

    public List<Map> lineCharts() {
        return userMapper.lineCharts();
    }

    public List<Map<String,Object>> pieCharts() {
        //最终的结果数据list
        List<Map<String,Object>> resultMapList = new ArrayList<Map<String,Object>>();
        //数据的构造如下
          /*[
            {
              name: '山东省',
              y: 7,
              drilldown: '山东省',
              id: '山东省'
              data: [
                {'济南市',5},
                {'威海市',5},
                {'青岛市',5}
              ]
            }
          ]*/

        //查询用户所有的数据
        List<User> userList = userMapper.selectAll();
        //根据省份分组,使用流式方程
        Map<String, List<User>> provinceMap = userList.stream().collect(Collectors.groupingBy(User::getProvince));
        //遍历集合构造结果数据
        for (String province : provinceMap.keySet()) {
            Map<String,Object> resultMap = new HashMap<>();
            resultMap.put("name",province);
            resultMap.put("drilldown",province);
            resultMap.put("id",province);
            //当前省份下的所有用户
            List<User> cityUserList = provinceMap.get(province);
            resultMap.put("y",cityUserList.size());

            //根据省份下面的城市进行分组
            Map<String, List<User>> cityMap = cityUserList.stream().collect(Collectors.groupingBy(User::getCity));
            List<Map<String,Object>> cityMapList = new ArrayList<>();
            for (String cityName : cityMap.keySet()) {
                Map<String,Object> dataMap = new HashMap<>();
                dataMap.put("name",cityName);
                dataMap.put("y",cityMap.get(cityName).size());cityMapList.add(dataMap);

            }

            resultMap.put("data",cityMapList);

            resultMapList.add(resultMap);
        }

        return resultMapList;
    }

    public Map<String,Object> pieECharts() {
        //返回的数据格式中市的值加起来要等于市的值，并且顺序要对应
        //返回数据格式：{"province" : [{name : "山东省"，value : 5}，{xxx}], "city" : [{name:"青岛市"，value:5}，xxx]}
        //构造返回的数据的map
        Map<String,Object> resultMap = new HashMap<String,Object>();

        List<Map<String,Object>> provinceMapList = new ArrayList<>();
        List<Map<String,Object>> cityMapList = new ArrayList<>();
        //有顺序根据省份和城市排序，条件查询
        Example example = new Example(User.class);
        example.setOrderByClause("province,city");
        List<User> userList = userMapper.selectByExample(example);

        //给查询的数据进行按省份分组，流式编程,分组时得排序，不然数据会乱
        Map<String, List<User>> provinceMap = userList.stream().collect(Collectors.groupingBy(User::getProvince, LinkedHashMap::new, Collectors.toList()));

        //将分组的省份，放到List集合中
        for (String province : provinceMap.keySet()) {
            Map<String,Object> map = new HashMap<String,Object>();
            map.put("name",province);
            map.put("value",provinceMap.get(province).size());
            provinceMapList.add(map);
        }

        //给查询的数据进行按城市分组，流式编程,分组时得排序，不然数据会乱
        Map<String, List<User>> cityMap = userList.stream().collect(Collectors.groupingBy(User::getCity, LinkedHashMap::new, Collectors.toList()));
        //将分组的城市，放到List集合中
        for (String city : cityMap.keySet()) {
            Map<String,Object> map = new HashMap<String,Object>();
            map.put("name",city);
            map.put("value",cityMap.get(city).size());
            cityMapList.add(map);
        }

        //将组装的数据放到最后的结果集中
        resultMap.put("province",provinceMapList);
        resultMap.put("city",cityMapList);
        return resultMap;
    }


}

package com.veo;

import com.veo.pojo.User;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler;
import org.apache.poi.xssf.usermodel.XSSFComment;

public class SheetHandler implements XSSFSheetXMLHandler.SheetContentsHandler {
    private User user= null;
    /**
     * 处理开始的行
     * @param rowIndex 行索引
     */
    @Override
    public void startRow(int rowIndex) {
        //跳过标题行，处理第二行数据行
        if (rowIndex == 0){
            user = null;
        }else{
            user = new User();
        }

    }

    /**
     * 处理每一行的所有单元格
     * @param cellName 单元格的名字
     * @param cellValue 单元格的值
     * @param xssfComment
     */
    @Override
    public void cell(String cellName, String cellValue, XSSFComment xssfComment) {
        if (user != null){
            //获取单元格的名字，A,B,C,D,E,F
            String letter = cellName.substring(0, 1);
            switch (letter){
                case "A":{
                    user.setId(Long.parseLong(cellValue));
                }
                case "B":{
                    user.setUserName(cellValue);
                }
            }
        }
    }

    /**
     * 处理每一行的结束
     * @param rowIndex 行索引
     */
    @Override
    public void endRow(int rowIndex) {
        if (rowIndex != 0){
            System.out.println(user);
        }
    }
}

package com.veo.utils;

import org.apache.commons.lang.StringUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.imageio.ImageIO;
import java.awt.image.BufferedImage;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.util.Map;

public class ExcelExportEngine {

    /**
     *
     * @param object 数据
     * @param workbook 工作薄
     * @return Workbook
     */
    public static Workbook writeToExcel(Object object,Workbook workbook,String imagePath) throws Exception {
        //使用将bean对象转换为一个map
        Map<String, Object> map = EntityUtils.entityToMap(object);

        //获取工作表
        Sheet sheet = workbook.getSheetAt(0);
        //循环100行，每一行循环100个单元格，获取为null就退出循环
        Row row = null;
        Cell cell = null;
        for (int i = 0; i < 100; i++) {
            row = sheet.getRow(i);
            if (row == null){
                break;
            }else {
                for (int j = 0; j < 100; j++) {
                    cell = row.getCell(j);
                    //解决中途单元格为null的情况数据不渲染
                    if (cell != null){
                        writeToCell(cell,map);
                    }
                }
            }
        }

        //图片的处理
        //图片路径不为空是才操作，复用代码，同时操作带和不带图片的文件导出
        if (imagePath != null){
            // 照片：第2行到第5行，第3列到第4列
            //创建字节输出流
            ByteArrayOutputStream byteArrayOutputStream = new ByteArrayOutputStream();
            //读取图片，将照片放入到了一个带缓存的类中
            BufferedImage bufferedImage = ImageIO.read(new File(imagePath));

            //获取文件的扩展名
            String extName = imagePath.substring(imagePath.lastIndexOf(".") + 1).toUpperCase();
            System.out.println(extName);
            //将文件写出到字节输出流中，参数缓存图片类，文件的格式，流
            ImageIO.write(bufferedImage,extName,byteArrayOutputStream);

            //Patriarch类控制图片的写入，ClientAnchor，控制照片的位置
            Drawing patriarch = sheet.createDrawingPatriarch();
            //获取模板引擎的第二张表来得到图片要插入的位置
            Sheet sheet1 = workbook.getSheetAt(1);
            int col1 = ((Double)sheet1.getRow(0).getCell(0).getNumericCellValue()).intValue();
            int row1 = ((Double)sheet1.getRow(0).getCell(1).getNumericCellValue()).intValue();
            int col2 = ((Double)sheet1.getRow(0).getCell(2).getNumericCellValue()).intValue();
            int row2 = ((Double)sheet1.getRow(0).getCell(3).getNumericCellValue()).intValue();

            //指定图片的位置：第2行到第5行，第3列到第4列
            //参数：0 左上角x轴的偏移量, 0 左上角y轴的偏移量, 0 右下角x轴的偏移量, 0 右下角y轴的偏移量, 0 起始列, 0 起始行, 0 结束列, 0 结束行
            //图片在结束列和结束行的左上角位置，要填充满的话要加 1，偏移单位1厘米=360000
            XSSFClientAnchor anchor = new XSSFClientAnchor(0, 0, 0, 0, col1, row1, col2, row2);

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
            //删除导出后的第二张图片定位表
            workbook.removeSheetAt(1);
        }

        return workbook;
    }

    /**
     * 比较单元格的键和map中定义的键是否相等，相等就使用map中的值替换单元格中的内容
     * @param cell 单元格
     * @param map map单个用户数据
     */
    private static void writeToCell(Cell cell, Map<String, Object> map) {
        //判断单元格是否为公式，是公式就不出来
        CellType cellType = cell.getCellType();
        switch (cellType){
            case FORMULA:{
                break;
            }
            default:{
                String cellValue = cell.getStringCellValue();
                if (StringUtils.isNotBlank(cellValue)){
                    for (String key : map.keySet()){
                        if (key.equals(cellValue)){
                            cell.setCellValue(map.get(key).toString());
                        }
                    }
                }
            }
        }
    }


}

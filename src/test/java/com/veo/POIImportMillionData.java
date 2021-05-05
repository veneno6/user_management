package com.veo;

//导入百万数据
public class POIImportMillionData {
    public static void main(String[] args) throws Exception {
        //下面默认使用dom4j读取Excel文件会导致内存溢出的异常
//        XSSFWorkbook workbook = new XSSFWorkbook("百万数据.xlsx");
//        XSSFSheet sheet = workbook.getSheetAt(0);
//        String stringCellValue = sheet.getRow(0).getCell(0).getStringCellValue();
//        System.out.println(stringCellValue);

        //要使用Sax方式读取Excel文件
        new ExcelParse().parse("百万数据.xlsx");
    }
}

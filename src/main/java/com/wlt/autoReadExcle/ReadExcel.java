package com.wlt.autoReadExcle;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.ArrayList;
import java.util.List;

/**
 * Created by One on 2016/8/19.
 */
public class ReadExcel {
    //获得文件的名称，用这个名字来作为数据库的名字和表的名字
    public static String getFileName(String path){
        File tempFile =new File(path.trim());
        return tempFile.getName();
    }
    //得到一个XSSWorkbook类型的数据
    public static XSSFWorkbook getXSSFWorkbook(String path) {
        //声明输入流
        InputStream is = null;
        try {
            //打开文件流
            is = new FileInputStream(path);
            //读取文件流到XSSWorkbook类型
            return new XSSFWorkbook(is);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
        return null;
    }
    public static XSSFSheet getSheet(XSSFWorkbook xssfWorkbook){

        return null;
    }
    //读取Excle表的字段，即就是得到所有的列，用这些名字来创建表字段
    public List<String> readExcelColumn(XSSFSheet xssfSheet) throws IOException {
        /*//得到XSSWorkbook类型文件流
        XSSFWorkbook xssfWorkbook = getXSSFWorkbook(path);
        //读取第一张表
        XSSFSheet xssfSheet = xssfWorkbook.getSheetAt(0);
        System.out.println(xssfSheet.getSheetName());*/
        //读取第一行数据
        XSSFRow row = xssfSheet.getRow(0);
        System.out.println("xssfSheet.getSheetName() : "+xssfSheet.getSheetName());
        //声明一个表格单元的数组
        List<String> cellList = new ArrayList<String>();
        //获取表的列数
        int columnNum = row.getLastCellNum();
        //保存每列的名称
        for (int i = 0; i < columnNum; i++) {
            //得到列名
            cellList.add(String.valueOf(row.getCell(i).getStringCellValue()));
        }
        return cellList;
    }
    //读取Excle中的数据
    public ArrayList readExcleData( XSSFSheet xssfSheet) {
        /*//得到XSSWorkbook类型文件流
        XSSFWorkbook xssfWorkbook = getXSSFWorkbook(path);
        //读取第一张表
        XSSFSheet xssfSheet = null;*/

        // 读取表内容，以表的数量做为循环次数限制
        //for (int numSheet = 0; numSheet < xssfWorkbook.getNumberOfSheets(); numSheet++) {
            //读取整张表
           /* System.out.println(numSheet);
            xssfSheet = xssfWorkbook.getSheetAt(numSheet);*/
            //判断是否含有数据
            if (xssfSheet == null) {
                return null;
            }
        //因为表里面可能不止一行数据，所以初始化一个存放List类型数据的List，来保存每一行的数据集合
        ArrayList<List> listcolumn = new ArrayList<List>();
            //循环读取每一行的数据
            for (int rowNum = 1; rowNum <= xssfSheet.getLastRowNum(); rowNum++) {
                //获得整行数据
                XSSFRow xssfRow = xssfSheet.getRow(rowNum);
                //初始化一个String类型的List，来保存每行的数据
                List<String> listData = new ArrayList<String>();
                //行元素判空
                if (xssfRow != null) {
                    //循环读取每一个表格里的数据
                    for (int column = 0; column < xssfRow.getLastCellNum(); column++) {
                        XSSFCell data = xssfRow.getCell(column);
                        //转换成String类型，保存进list集合
                        listData.add(String.valueOf(data));
                    }
                    //每行的数据集合，添加到表的list集合里面
                    listcolumn.add(listData);
                    System.out.println(listcolumn);
                }
            }
            return listcolumn;
            //System.out.println(listcolumn);
        //}
        //return null;
    }
}

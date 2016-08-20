package com.wlt.autoReadExcle;

import com.oracle.deploy.update.UpdateInfo;

import java.io.IOException;
import java.io.InputStream;
import java.sql.*;
import java.util.ArrayList;
import java.util.List;
import java.util.Properties;

/**
 * Created by One on 2016/8/19.
 */
public class CreateDatabase {
    //声明几个全局变量
    private static String name="";
    private static String driverclassname;
    private static String url;
    private static String username;
    private static String password;
    //静态语句块，加载properties文件，并给变量赋值
    static {
        try {
            InputStream in = UpdateInfo.class.getClassLoader().getResourceAsStream("db.properties");
            Properties props = new Properties();
            props.load(in);
            driverclassname = props.getProperty("driverclassname");
            url = props.getProperty("url");
            username = props.getProperty("user");
            password = props.getProperty("password");
        } catch (IOException e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        }
    }

    /**
     * 这个方法用来创建数据库和表。
     * 首先说一下这里的大概思路，我们在创建自己新的数据库之前，得先链接上数据库，
     * 所以我们可以先链接任一数据库文件，然后再来创建自己的数据库。接着创建表和字段。
     * @param name
     * @param columnList
     */
    public static void createDataBase(String name, List columnList) {
        // 创建一个数据库连接
        Connection con = null;
        // 创建预编译语句对象，一般都是用这个而不用Statement
        PreparedStatement pre = null;
        try {
            //强制JVM将com.mysql.jdbc.Driver这个类加载入内存，并将其注册到DriverManager类
            Class.forName(driverclassname);
            // 获取连接
            con = DriverManager.getConnection(url+"test", username, password);
            /**
             * 创建表的说起来语句，但是这里要注意的是要注意在必要的地方加上空格，
             * 以此保证sql语句可以通过编译，还有一个麻烦的地方就是如果要创建主键的话会稍微有点麻烦。
             * 还有就是所有字段都是一个长度，会造成空间浪费
             * 这里没有写，业务需要的时候，自己注意修改
             */
            String createTable="create table "+name+" (";
            String limit= " varchar(100),";
            String end=");";
            for (Object columnName : columnList){
                createTable+=columnName+" "+ limit;
            }
            /**
             * 因为上面的语句给每个字段后面都即使有逗号，但是在最后一个字段不需要逗号，要对它进行删除操作。
             */
            createTable=createTable.substring(0,createTable.length()-1);
            createTable+=end;
            /**
             * 创建数据库的sql语句，这个比创建表的语句要简单点。
             */
            String databaseSql = "create database " + name+";";
            /**
             * 最好对sql语句进行预编译，这样可以有效防止sql注入。
             */
            pre = con.prepareStatement(databaseSql);
            pre.executeUpdate(databaseSql);
            pre.close();
            con.close();
            /**
             * 这里链接刚才创建的数据库，后面才能在新的数据库中创建表
             */
            Connection newConn = DriverManager.getConnection("jdbc:mysql://localhost/" + name,username, password);
            if (newConn!=null){
                pre =newConn.prepareStatement(createTable);
                pre.executeUpdate(createTable);
                pre.close();
                newConn.close();
            }
        } catch (ClassNotFoundException e) {
            e.printStackTrace();
        } catch (SQLException e) {
            e.printStackTrace();
        }
    }

    /**
     * 将读取到的Excle表中的数据插入到数据库中。许多地方与上面类似，不再赘述。提醒一下，在创建sql语句时一定要注意格式。
     * @param listcolumn
     */
    public static void insertData(ArrayList<List> listcolumn){
        Connection con = null;
        PreparedStatement pre = null;
        try {
            Class.forName(driverclassname);//
            con = DriverManager.getConnection(url+name, username, password);// 获取连接
            for (List column:listcolumn){
                String dataSql="insert into "+name +" values (";
                for (Object data:column){
                    dataSql+="'"+data.toString()+"'"+",";
                }
                dataSql=dataSql.substring(0,dataSql.length()-1);
                dataSql+=");";
                pre = con.prepareStatement(dataSql);
                pre.executeUpdate(dataSql);
            }
            pre.close();
            con.close();
        } catch (ClassNotFoundException e) {
            e.printStackTrace();
        } catch (SQLException e) {
            e.printStackTrace();
        }
    }

    /**
     * 下面写一个主方法，对以上的方法进行测试
     * @param args
     */
    public static void main(String[] args) {
        String file="E://readExcle.xlsx";
        List list = null;
        ReadExcel readExcel = new ReadExcel();
        name = readExcel.getFileName(file);
        name=name.substring(0,name.indexOf("."));
        try {
            list = readExcel.readExcelColumn(file);
        } catch (IOException e) {
            e.printStackTrace();
        }
        createDataBase(name,list);
        ArrayList<List> listcolumn= readExcel.readExcleData(file);
        insertData(listcolumn);
    }
}
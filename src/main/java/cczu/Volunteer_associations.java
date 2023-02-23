package cczu;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;

import java.io.FileOutputStream;
import java.io.IOException;

//青年志愿者协会
public class Volunteer_associations {
    private static final String afterPATH = "生成后的空课表/";
    public static final String beforePATH = "原始空课表/";
    public static final String middlePATH = "工程文件/";
    public static final String[] Department;
    static {
         Department = new String[]{"行政部", "宣传部", "办公室","人事部","项目部","新闻部","外联部"};
     }

     private Volunteer_associations(){}

    //初始化
     public static void Init() throws IOException {
         //空课表模版初始化
         for (String str : Department) {
             init(str);
         }
         //工程文件初始化
         engineeringFile();
     }

    //生成课程表模版
    private static void init(String department) throws IOException {
        //创建一个工作簿
        Workbook workbook = new HSSFWorkbook();
        CellStyle cellStyle = workbook.createCellStyle();
        cellStyle.setWrapText(true);
        //创建单周工作表
        Sheet Single_week = workbook.createSheet("单周");
        //创建双周工作表
        Sheet Double_week = workbook.createSheet("双周");

        for (int j = 0; j < 10; j++) {
            if (j == 0){
                //创建一个行
                Row row = Single_week.createRow(0);
                Row row1 = Double_week.createRow(0);
                //创建第一行数据
                String[] arr = {"一","二","三","四","五"};
                for (int i = 0;i <=4 ; i++){
                    Cell cell = row.createCell(i+1);
                    Cell cell1 = row1.createCell(i + 1);
                    cell.setCellValue("周"+arr[i]);
                    cell1.setCellValue("周"+arr[i]);
                }
            }else {
                Row rows = Single_week.createRow(j);
                Row rows1 = Double_week.createRow(j);
                //创建第一列数据
                Cell cell = rows.createCell(0);
                Cell cell1 = rows1.createCell(0);
                cell.setCellValue("第"+j+"节课");
                cell1.setCellValue("第"+j+"节课");
                //创建第二列到后的单元格
                for (int i = 1; i < 6; i++) {
                    Cell cell2 = rows.createCell(i);
                    cell2.setCellValue(" ");
                    cell2.setCellStyle(cellStyle);
                    Cell cell3 = rows1.createCell(i);
                    cell3.setCellValue(" ");
                    cell3.setCellStyle(cellStyle);
                }
            }
        }


        //生成一张表(IO流)，03版本就是使用xls结尾
        FileOutputStream fileOutputStream = new FileOutputStream(afterPATH + String.format("%s空课表.xls", department));
        //输出
        workbook.write(fileOutputStream);
        //关闭IO流
        fileOutputStream.close();
    }

    //工程文件模版
    private static void engineeringFile() throws IOException {
        //创建一个工作簿
        Workbook workbook = new HSSFWorkbook();
        //创建工作表
        for (int i = 0, departmentLength = Department.length; i < departmentLength+1; i++) {
            if (i == departmentLength){
                Sheet sheet = workbook.createSheet("工程文件");
            }else {
                String s = Department[i];
                Sheet sheet = workbook.createSheet(s);
            }
        }
        FileOutputStream fileOutputStream = new FileOutputStream(middlePATH+"工程文件.xls");
        workbook.write(fileOutputStream);
        fileOutputStream.close();
    }
}

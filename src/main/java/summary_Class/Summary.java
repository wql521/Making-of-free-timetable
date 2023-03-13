package summary_Class;

import cczu.TextToVoice;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;

import javax.sound.sampled.LineUnavailableException;
import javax.sound.sampled.UnsupportedAudioFileException;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

public class Summary {
    static String[] volunteer_departments = {"行政部", "宣传部", "办公室","人事部","项目部","新闻部","外联部"};
    static final String afterPATH = "生成后的空课表/";
    static final String middlePATH = "工程文件/";
    public static void main(String[] args) throws IOException, UnsupportedAudioFileException, LineUnavailableException {
        init();
        for (int i = 0; i < volunteer_departments.length; i++) {
            summary_class(volunteer_departments[i]);
        }
        System.out.println("全部汇总完成");
        TextToVoice.Path = "/Users/wangqianlong/Desktop/All/IntelliJ IDEA/语言合成_wav/";
        TextToVoice.Text_Voice("全部汇总完成");
    }

    public static void write(String name,int sheet,int row,int cell) throws IOException {
        //获取数据流
        FileInputStream fileInputStream = new FileInputStream(middlePATH+"汇总空课表.xls");
        //获取工作簿
        HSSFWorkbook workbook = new HSSFWorkbook(fileInputStream);
        if (sheet == 0){
            //获取工作表
            HSSFSheet sheetAt = workbook.getSheet("单周");
            //获取行
            HSSFRow row1 = sheetAt.getRow(row);
            //获取单元格
            HSSFCell cell1 = row1.getCell(cell);
            //先获取数据，再添加数据
            String stringCellValue = cell1.getStringCellValue();
            cell1.setCellValue(stringCellValue+" "+name);

            //输出文件流
            FileOutputStream fileOutputStream = new FileOutputStream(middlePATH + "汇总空课表.xls");
            //写入文件流
            workbook.write(fileOutputStream);
            //关闭文件流
            fileOutputStream.close();
            fileInputStream.close();

        } else if (sheet == 1) {
            //获取工作表
            HSSFSheet sheetAt = workbook.getSheet("双周");
            //获取行
            HSSFRow row1 = sheetAt.getRow(row);
            //获取单元格
            HSSFCell cell1 = row1.getCell(cell);
            //先获取数据，再添加数据
            String stringCellValue = cell1.getStringCellValue();
            cell1.setCellValue(stringCellValue+" "+name);

            //输出文件流
            FileOutputStream fileOutputStream = new FileOutputStream(middlePATH + "汇总空课表.xls");
            //写入文件流
            workbook.write(fileOutputStream);
            //关闭文件流
            fileOutputStream.close();
            fileInputStream.close();
        }
    }

    public static void summary_class(String department) throws IOException {
        //获取数据流
        FileInputStream fileInputStream = new FileInputStream(afterPATH + department + "空课表.xls");
        //获取一个工作簿
        HSSFWorkbook workbook = new HSSFWorkbook(fileInputStream);
        //获取工作表
        //单周
        HSSFSheet Single_week = workbook.getSheet("单周");
        //双周
        HSSFSheet Bi_weekly = workbook.getSheet("双周");
        //获取行数
        //单周
        int physicalNumberOfRows_single = Single_week.getPhysicalNumberOfRows();
        //双周
        int physicalNumberOfRows_double = Bi_weekly.getPhysicalNumberOfRows();

        //单周
        for (int i = 1; i < physicalNumberOfRows_single; i++) {
            //获取行
            HSSFRow row = Single_week.getRow(i);
            //获取单元格内容
            for (int j = 1; j < 6; j++) {
                HSSFCell cell = row.getCell(j);
                if (cell != null){
                    String stringCellValue = cell.getStringCellValue();
                    //System.out.println(stringCellValue);
                    write(stringCellValue,0,i,j);
                }
            }
        }

        //双周
        for (int i = 1; i < physicalNumberOfRows_double; i++) {
            //获取行
            HSSFRow row = Bi_weekly.getRow(i);
            //获取单元格内容
            for (int j = 1; j < 6; j++) {
                HSSFCell cell = row.getCell(j);
                if (cell != null){
                    String stringCellValue = cell.getStringCellValue();
                    System.out.println(stringCellValue);
                    write(stringCellValue,1,i,j);
                }
            }
        }
        fileInputStream.close();
    }

    public static void init() throws IOException {
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
        FileOutputStream fileOutputStream = new FileOutputStream(middlePATH + "汇总空课表.xls");
        //输出
        workbook.write(fileOutputStream);
        //关闭IO流
        fileOutputStream.close();
    }
}

package analyse;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class Analyse {
    static final String middlePATH = "工程文件/";
    public static void main(String[] args) throws IOException {
        read();
    }

    public static void read() throws IOException {
        //获取数据流
        FileInputStream fileInputStream = new FileInputStream(middlePATH + "汇总空课表.xls");
        //获取工作簿
        HSSFWorkbook workbook = new HSSFWorkbook(fileInputStream);
        //获取工作表
        //单周
        HSSFSheet Single_week = workbook.getSheetAt(0);
        //双周
        HSSFSheet Bi_weekly = workbook.getSheetAt(1);
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
                String stringCellValue = cell.getStringCellValue();
                int rec = judge(stringCellValue);
                HSSFCell cell1 = row.createCell(j + 7);
                cell1.setCellValue(rec);
            }
        }

        //双周
        for (int i = 1; i < physicalNumberOfRows_double; i++) {
            //获取行
            HSSFRow row = Bi_weekly.getRow(i);
            //获取单元格内容
            for (int j = 1; j < 6; j++) {
                HSSFCell cell = row.getCell(j);
                String stringCellValue = cell.getStringCellValue();
                int rec = judge(stringCellValue);
                HSSFCell cell1 = row.createCell(j + 7);
                cell1.setCellValue(rec);
            }
        }
        //输出文件流
        FileOutputStream fileOutputStream = new FileOutputStream(middlePATH + "汇总空课表.xls");
        //写入文件流
        workbook.write(fileOutputStream);
        //关闭文件流
        fileOutputStream.close();
        fileInputStream.close();

    }

    public static int judge(String s){
        String pattern = "([\\u4e00-\\u9fa5][\\u4e00-\\u9fa5][\\u4e00-\\u9fa5])|([\\u4e00-\\u9fa5][\\u4e00-\\u9fa5])";
        Pattern r = Pattern.compile(pattern);
        int count = 0;
        Matcher m = r.matcher(s);
        while (m.find()){
            count = count+1;
            //System.out.println(m.group());
        }
        System.out.println(count);
        return count;
    }
}

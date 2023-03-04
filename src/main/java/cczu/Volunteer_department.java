package cczu;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import javax.sound.sampled.LineUnavailableException;
import javax.sound.sampled.UnsupportedAudioFileException;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;

//志愿者部门
public class Volunteer_department {
    //存储一节课第几周开始
    //static ArrayList<String> min_Class = new ArrayList<>();
    //存储一节课第几周结束
    //static ArrayList<String> max_Class = new ArrayList<>();

    //构造方法私有化
    private Volunteer_department(){}

    //获取对应空课表文件
    public static void fetchFiles() throws IOException, UnsupportedAudioFileException, LineUnavailableException {
        //加载课表文件
        Volunteer.listFilesInit();
        for (int i = 0; i < 7; i++) {
            for (int j = 0; j < Volunteer.arrayList.get(i).size()-1; j++) {
                //获取单个课表文件名字
                String filePath = Volunteer_associations.beforePATH
                        +Volunteer.arrayList.get(i).get(Volunteer.arrayList.get(i).size()-1).toString()
                        +"/"
                        +Volunteer.arrayList.get(i).get(j).toString();
                writeFiles(filePath,volunteerUtil.getdepartment_filePath(filePath));
            }
        }
    }

    //写入对应空课表
    private static void writeFiles(String filePath,String department) throws IOException {
        //泛型集合，存储了个人一周的课表信息
        ArrayList<ArrayList> arrayList = volunteerUtil.openFiles(filePath);
        int day = arrayList.size();//5
        int lesson = 9;
        String Name = volunteerUtil.getName_filePath(filePath);
        //打开工作流
        String Path = "生成后的空课表/"+department+"空课表.xls";
        FileInputStream fileInputStream = new FileInputStream(Path);
        //打开工作簿
        HSSFWorkbook sheets = new HSSFWorkbook(fileInputStream);
        //打开工作表
        HSSFSheet Single_week = sheets.getSheet("单周");
        HSSFSheet Bi_weekly = sheets.getSheet("双周");

        for (int i = 0; i < day; i++) {
            ArrayList<String> datList = arrayList.get(i);

            for (int j = 0; j < lesson; j++) {
                //获取行
                HSSFRow Single_row = Single_week.getRow(j+1);
                HSSFRow bi_weekly_Row = Bi_weekly.getRow(j+1);

                //获取单元格
                HSSFCell single_rowCell = Single_row.getCell(i+1);
                HSSFCell bi_weekly_rowCell = bi_weekly_Row.getCell(i+1);

                //获取判断结果
                //在这里可以获得个人的课程信息
                String getLesson = datList.get(j);
                /* 获取一个课程的开始周和结束周
                if (getLesson.length()>0){
                    volunteerUtil.get_start_over(getLesson);
                }
                 */
                int src = volunteerUtil.judgementClass(getLesson);

                //写入对应单元格内容
                //都不写入为0 单周写入为1 双周写入为2 单双周都写入为3
                switch (src){
                    case 1 -> {
                        if (single_rowCell.getStringCellValue().equals(" ")){
                            single_rowCell.setCellValue(Name);
                        }else {
                            single_rowCell.setCellValue(single_rowCell.getStringCellValue()+" "+Name);
                            //bi_weekly_rowCell.setCellValue("");
                        }
                    }
                    case 2 -> {
                        if (bi_weekly_rowCell.getStringCellValue().equals(" ")){
                            bi_weekly_rowCell.setCellValue(Name);
                        }else {
                            //single_rowCell.setCellValue("");
                            bi_weekly_rowCell.setCellValue(bi_weekly_rowCell.getStringCellValue()+" "+Name);
                        }

                    }
                    case 3 -> {
                        if (single_rowCell.getStringCellValue().equals(" ") || bi_weekly_rowCell.getStringCellValue().equals(" ")){
                            single_rowCell.setCellValue(Name);
                            bi_weekly_rowCell.setCellValue(Name);
                        }else {
                            single_rowCell.setCellValue(single_rowCell.getStringCellValue()+" "+Name);
                            bi_weekly_rowCell.setCellValue(bi_weekly_rowCell.getStringCellValue()+" "+Name);
                        }
                    }
                    case 0 -> {
                    }
                }
            }

        }


        //输出文件流
        FileOutputStream fileOutputStream = new FileOutputStream(Path);
        //写入工作簿
        sheets.write(fileOutputStream);
        //关闭输出流
        fileOutputStream.close();
        //关闭上面输入文件流
        fileInputStream.close();
    }


}

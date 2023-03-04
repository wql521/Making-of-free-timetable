package cczu;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;

//工具类
public final class volunteerUtil {

    private volunteerUtil(){}

    /**
     *
     * @param list 传入集合
     * @param department 部门名称
     * @param Row 需要写入的目标行
     */
    public static void writeEngineering(ArrayList<String> list,String department,int Row) throws IOException {
        String filePATH = "工程文件/工程文件.xls";
        //获取文件流
        FileInputStream fileInputStream = new FileInputStream(filePATH);
        //获取一个工作簿
        HSSFWorkbook hssfWorkbook = new HSSFWorkbook(fileInputStream);
        //获取一个工作表
        HSSFSheet sheet = hssfWorkbook.getSheet("工程文件");
        //获取行
        HSSFRow row = sheet.createRow(Row);
        for (int i = 0; i < list.size()+1; i++) {
            //写入数据
            if (i == 0){
                HSSFCell cell = row.createCell(i);
                cell.setCellValue(department);
            }else {
                HSSFCell cell = row.createCell(i);
                cell.setCellValue(list.get(i - 1));
            }
        }
        //关闭io
        fileInputStream.close();

        //输出文件流
        FileOutputStream fileOutputStream = new FileOutputStream(filePATH);
        //写入
        hssfWorkbook.write(fileOutputStream);
        //关闭文件流
        fileOutputStream.close();
    }

    //都不写入为0 单周写入为1 双周写入为2 单双周都写入为3
    public static int judgementClass(String course){
        if (course.length() == 0){
            return 3;
        }else {
            boolean singleSrc = course.matches(".*\\s[单]\\s{1}.*");
            boolean doubleSrc = course.matches(".*\\s双\\s{1}.*");
            if (singleSrc && !doubleSrc){
                return 2;
            }else if (!singleSrc && doubleSrc){
                return 1;
            }else {
                return 0;
            }
        }
    }

    //打开课表文件
    public static ArrayList<ArrayList> openFiles(String filePath) throws IOException {
        ArrayList<ArrayList> arrayList = new ArrayList<>();
        //获取文件流
        FileInputStream fileInputStream = new FileInputStream(filePath);
        //打开工作簿
        HSSFWorkbook sheets = new HSSFWorkbook(fileInputStream);
        //打开工作表
        HSSFSheet sheetAt = sheets.getSheetAt(0);

        for (int j = 3 ;j <=7;j++) {
            //存储一天的课程信息
            ArrayList<String> list =new ArrayList<>();
            for (int i = 0; i < 18; i++) {
                if (i % 2 != 0){
                    //获取行
                    HSSFRow row = sheetAt.getRow(i);
                    //获取单元格
                    HSSFCell cell = row.getCell(j);
                    if (cell == null){
                        list.add("");
                    }else {
                        list.add(cell.getStringCellValue());
                    }
                }
            }
            arrayList.add(list);
            System.out.println(list.toString());
        }
        System.out.println(getName_filePath(filePath)+"课程表加载完毕");
        //关闭文件流
        fileInputStream.close();
        return arrayList;
    }

    //获取文件名中的个人姓名
    public static String getName_filePath(String filePath){
        String[] Temp = filePath.split("/");
        String[] temp = Temp[2].split("\\.");
        return temp[0];
    }
    //获取文件名中的部门名称
    public static String getdepartment_filePath(String filePath){
        String[] Temp = filePath.split("/");
        return Temp[1];
    }


    /* 判断开始周和结束周方法(采用正则表达式)
    //获取一节课第几周开始，第几周结束
    public static void get_start_over(String str){
        String pattern = "(1[0-8]|[0-9])-(1[0-8]|[0-9])";

        ArrayList<String> list = new ArrayList<>();

        Pattern r = Pattern.compile(pattern);
        Matcher m = r.matcher(str);
        while (m.find()){
            String s = m.group();
            list.add(s);
        }
        String min_src = list.get(0);
        String max_src = list.get(list.size()-1);
        //第几周开始，记录并且添加到Volunteer_department中的集合中
        char c = min_src.charAt(0);
        String s = String.valueOf(c);
        Volunteer_department.min_Class.add(s);
        //System.out.println(c);


        //用于获取最后一个数据段的列表
        ArrayList<String> max_List = new ArrayList<>();
        //对最大时间长度进行截取
        Pattern compile = Pattern.compile("(1[0-8]|[0-9])");
        Matcher matcher = compile.matcher(max_src);
        while (matcher.find()){
            String s1 = matcher.group();
            max_List.add(s1);
        }
        //第几周结束，记录并且添加到Volunteer_department中的集合中
        Volunteer_department.max_Class.add(max_List.get(1));
        //System.out.println(max_List.get(1));
    }

     */



}

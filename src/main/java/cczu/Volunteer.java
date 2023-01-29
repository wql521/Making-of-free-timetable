package cczu;

import java.io.File;
import java.io.FilenameFilter;
import java.io.IOException;
import java.util.ArrayList;

//志愿者
public class Volunteer{
    private Volunteer(){}

    //原始空课表文件夹下七个部门所有课表
    //后期需要对文件进行判断
    public static ArrayList<ArrayList> arrayList = new ArrayList<>();

    public static void listFilesInit() throws IOException {
        for (int i = 0; i < Volunteer_associations.Department.length; i++) {
            ArrayList<String> list = judgementFile(Volunteer_associations.Department[i]);
            //写入工程文件-工程
            //将原始空课表中各部门中获取的文件名，写入
            volunteerUtil.writeEngineering(list,Volunteer_associations.Department[i],i);
            //课表文件路径写入泛型集合
            arrayList.add(list);
        }
        System.out.println("课表文件加载完毕");
    }

    //判断并且获取文件夹下所有xls文件
    private static ArrayList<String> judgementFile(String department){
        File file = new File(Volunteer_associations.beforePATH + department);
        ArrayList<String> list = null;
        if (!file.exists()){
            //判断是否存在部门文件夹,若不存在开始创建
            file.mkdirs();
            System.out.println("创建"+department+"文件夹");
            System.out.println("请将"+department+"的课表文件放入"+department+"文件夹下");
        } else if (file.list().length == 0){
            System.out.println(department+"没有课表文件");
        }else {
            list = getFileNames(file);
            list.add(department);
            System.out.println(department+":"+list.toString());
        }
        return list;
    }
    private static ArrayList<String> getFileNames(File file){
        ArrayList<String> list = new ArrayList<>();
        String[] arrSt = file.list();
        for (int i = 0; i < arrSt.length; i++) {
            boolean src = new FilenameFilter(){

                @Override
                public boolean accept(File dir, String name) {
                    return name.endsWith("xls");
                }
            }.accept(file,arrSt[i]);
            if (src){
                list.add(arrSt[i]);
            }
        }
        return list;
    }
}

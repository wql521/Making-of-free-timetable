package cczu;

import javax.sound.sampled.LineUnavailableException;
import javax.sound.sampled.UnsupportedAudioFileException;
import java.io.IOException;

public class Main {
    public static void main(String[] args) throws IOException, UnsupportedAudioFileException, LineUnavailableException {
        TextToVoice.Path="/Users/wangqianlong/Desktop/All/IntelliJ IDEA/语言合成_wav/";
        //协会
        //初始化
        Volunteer_associations.Init();
        //部门
        //生成空课表
        Volunteer_department.fetchFiles();

        //播放语音
        String s = data_analysis();
        System.out.println(s);
        TextToVoice.Text_Voice(s);

        //开始周数据
        //System.out.println(Volunteer_department.min_Class.toString());
        //结束周数据
        //System.out.println(Volunteer_department.max_Class.toString());
        //数据分析
        /* 循环
        while (true){
            Scanner sc = new Scanner(System.in);
            System.out.print("请输入第几周开始课程:");
            String start = sc.next();
            System.out.print("请输入第几周结束课程:");
            String end = sc.next();
            data_Analysis(start,end);
        }
         */
    }

    /* 方法
    public static void data_Analysis(String start,String end){
        int min = 0; //第几周开始课程
        int max = 0; //第几周结束课程
        //遍历开始周数据
        for (int i = 0; i < Volunteer_department.min_Class.size(); i++) {
            if (Volunteer_department.min_Class.get(i).equals(start)){
                System.out.println(Volunteer_department.min_Class.get(i));
                min = min+1;
            }
        }
        for (int i = 0; i < Volunteer_department.max_Class.size(); i++) {
            if (Volunteer_department.max_Class.get(i).equals(end)){
                max = max+1;
            }
        }

        System.out.println("第"+start+"开始课程:"+min+" "+"总课程数:"+Volunteer_department.min_Class.size());
        System.out.println("第"+end+"周结束课程:"+max+" "+"总课程数:"+Volunteer_department.min_Class.size());

    }
     */

    public static String data_analysis(){
        String rec = "Hi，boss，空课表全部生成。本次成功处理课表为：";
        for (int i = 0; i < Volunteer.arrayList.size(); i++) {
            int Size = Volunteer.arrayList.get(i).size();
            int max = Size-1;
            rec = rec + Volunteer.arrayList.get(i).get(max) + max +"人" + ",";
        }
        return rec;
    }
}
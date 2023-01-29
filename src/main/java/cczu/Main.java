package cczu;

import java.io.IOException;

public class Main {
    public static void main(String[] args) throws IOException {
        //协会
        //初始化
        Volunteer_associations.Init();
        //部门
        //生成空课表
        Volunteer_department.fetchFiles();
    }
}
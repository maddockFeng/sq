package sq.coolvisit.com;


import com.alibaba.fastjson.JSON;
import com.alibaba.fastjson.JSONArray;
import com.alibaba.fastjson.JSONObject;
import jxl.Workbook;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;

import java.io.*;
import java.sql.Array;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class main {
    public static void main(String[] args) {
        String path = main.class.getClassLoader().getResource("fk.json").getPath();
        String s = readJsonFile(path);
        JSONArray ja = JSON.parseArray(s);
        List<Map> mapList = new ArrayList<Map>();
        for(int i=0 ;i<ja.size();i++){
            JSONObject jb = ja.getJSONObject(i);
            String name = jb.getString("name");
            String phone = jb.getString("phone");
            String avatar = jb.getString("avatar");
            if (phone == null || phone.equals("")){
                if(i<10){
                    phone = "1390000000"+String.valueOf( i);
                }else
                if(i<100){
                    phone = "139000000"+String.valueOf( i);
                }else {
                    phone = "13900000"+String.valueOf( i);
                }
            }
            String email = "34"+i+"@qq.com";
            if (avatar == null || avatar.equals("")) {
                JSONArray photos = jb.getJSONArray("photos");
                if (photos.size() > 0) {
                    avatar = photos.getString(0);
                }
            }
            FixFileName("F:/fk/"+avatar,phone);
            System.out.println("index:"+i+" "+name+" " + phone+" "+avatar);
            HashMap<String,String> map = new HashMap<String, String>();
            map.put("name",name);
            map.put("phone",phone);
            map.put("email",email);
            mapList.add(map);

        }


        writeExcel(mapList);
    }

    /**
     * 读取json文件，返回json串
     * @param fileName
     * @return
     */
    public static String readJsonFile(String fileName) {
        String jsonStr = "";
        try {
            File jsonFile = new File(fileName);
            FileReader fileReader = new FileReader(jsonFile);

            Reader reader = new InputStreamReader(new FileInputStream(jsonFile),"utf-8");
            int ch = 0;
            StringBuffer sb = new StringBuffer();
            while ((ch = reader.read()) != -1) {
                sb.append((char) ch);
            }
            fileReader.close();
            reader.close();
            jsonStr = sb.toString();
            return jsonStr;
        } catch (IOException e) {
            e.printStackTrace();
            return null;
        }
    }

    /**
     * 通过文件路径直接修改文件名
     *
     * @param filePath    需要修改的文件的完整路径
     * @param newFileName 需要修改的文件的名称
     * @return
     */
    private static String FixFileName(String filePath, String newFileName) {
        File f = new File(filePath);
        if (!f.exists()) { // 判断原文件是否存在（防止文件名冲突）
            return null;
        }
        newFileName = newFileName.trim();
        if ("".equals(newFileName) || newFileName == null) // 文件名不能为空
            return null;
        String newFilePath = null;
        if (f.isDirectory()) { // 判断是否为文件夹
            newFilePath = filePath.substring(0, filePath.lastIndexOf("/")) + "/" + newFileName;
        } else {
            newFilePath =  "F:/fk/" + newFileName
                    + filePath.substring(filePath.lastIndexOf("."));
        }
        File nf = new File(newFilePath);
        try {
            f.renameTo(nf); // 修改文件名
        } catch (Exception err) {
            err.printStackTrace();
            return null;
        }
        return newFilePath;
    }


    public static void writeExcel(List<Map> maplist){
                //开始写入excel,创建模型文件头
                String[] titleA = {"name","phone","email","nickname","dp","remark"};
                //创建Excel文件，B库CD表文件
                File fileA = new File("F:/TestFile.xls");
                if(fileA.exists()){
                    //如果文件存在就删除
                    fileA.delete();
                }
                try {
                    fileA.createNewFile();
                    //创建工作簿
                    WritableWorkbook workbookA = Workbook.createWorkbook(fileA);
                    //创建sheet
                    WritableSheet sheetA = workbookA.createSheet("sheet1", 0);
                    Label labelA = null;
                    //设置列名
                    for (int i = 0; i < titleA.length; i++) {
                        labelA = new Label(i,0,titleA[i]);
                        sheetA.addCell(labelA);    
                    }            
                    //获取数据源
                    for (int i = 0; i < maplist.size(); i++) {
                        Map map = maplist.get(i);
                        labelA = new Label(0,i, (String) map.get("name"));
                        sheetA.addCell(labelA);
                        labelA = new Label(1,i,(String) map.get("phone"));
                        sheetA.addCell(labelA);
                        labelA = new Label(2,i,(String) map.get("email"));
                        sheetA.addCell(labelA);
                    }
                    workbookA.write();    //写入数据        
                    workbookA.close();  //关闭连接
                    System.out.println("成功写入文件，请前往E盘查看文件！");
 
                } catch (Exception e) {
                    System.out.println("文件写入失败，报异常...");
                }
    }

}

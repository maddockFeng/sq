package sq.coolvisit.com;


import com.alibaba.fastjson.JSON;
import com.alibaba.fastjson.JSONArray;
import com.alibaba.fastjson.JSONObject;
import jxl.Workbook;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;

import java.io.*;
import java.net.HttpURLConnection;
import java.net.URL;
import java.util.*;

import static java.lang.Thread.sleep;

/**
 * 导入只有照片的文件格式到excel表，并将照片改名为工号
 */
public class mainCX {
    static String outPath = "f://avatarFile/";
    public static void main(String[] args) {
        traverseFolder1("f://in");


    }

    public static void addPerson(String path,int index){
        String imageUrl = uploadFile("http://192.168.0.10:8888/common/upload",path);
        String name = path.substring(path.lastIndexOf("/")+1, path.lastIndexOf(".")) ;

        if (imageUrl == null || imageUrl.equals("")){
            System.out.println("update image error:"+path);
            return;
        }
        JSONObject jb = new JSONObject();
        jb.put("auth",1);
        jb.put("avatarUrl",imageUrl);
        jb.put("name",name);
        jb.put("personId","100000"+index);
        jb.put("status","0");
        jb.put("type","0");
        post("http://192.168.0.10:8888/person/create",jb.toString());
    }

    public static void traverseFolder1(String path) {
        int fileNum = 0, folderNum = 0;
        File file = new File(path);
        if (file.exists()) {
            LinkedList<File> list = new LinkedList<File>();
            LinkedList<File> listFile = new LinkedList<File>();
            File[] files = file.listFiles();
            for (File file2 : files) {
                if (file2.isDirectory()) {
                    System.out.println("文件夹:" + file2.getAbsolutePath());
                    list.add(file2);
                    folderNum++;
                } else {
                    System.out.println("文件:" + file2.getAbsolutePath());
                    listFile.add(file2);
                    fileNum++;
                }
            }
            File temp_file;
            while (!list.isEmpty()) {
                temp_file = list.removeFirst();
                files = temp_file.listFiles();
                for (File file2 : files) {
                    if (file2.isDirectory()) {
                        System.out.println("文件夹:" + file2.getAbsolutePath());
                        list.add(file2);
                        folderNum++;
                    } else {
                        System.out.println("文件:" + file2.getAbsolutePath());
                        listFile.add(file2);
                        fileNum++;
                    }
                }
            }

                //TODO 调用接口直接上传

//            for(int i=0;i<listFile.size();i++){
//                System.out.println("add person "+i);
//                addPerson(listFile.get(i).getPath(),i);
//                try {
//                    sleep(1000*10);
//                } catch (InterruptedException e) {
//                    e.printStackTrace();
//                }
//
//            }


            List<Map> mapList = new ArrayList<Map>();
            //生成excel表
            for(int i=0;i<listFile.size();i++){
                Map map = new HashMap();
                File f = listFile.get(i);
                map.put("name",f.getName().substring(0,f.getName().lastIndexOf(".")));
                map.put("phone","1000"+i);
                mapList.add(map);

                //文件改名
                FixFileName(f.getPath(), (String) map.get("phone"));
            }

            writeExcel(mapList);
        } else {
            System.out.println("文件不存在!");
        }
        System.out.println("文件夹共有:" + folderNum + ",文件共有:" + fileNum);

    }

    public void import1(){
        String path = mainCX.class.getClassLoader().getResource("fk.json").getPath();
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
            newFilePath =  outPath + newFileName
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
                File fileA = new File(outPath+"csvFile.xls");
                System.out.println("人数:"+maplist.size());
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
                        labelA = new Label(0,i,(String) map.get("phone"));
                        sheetA.addCell(labelA);

                        labelA = new Label(1,i, (String) map.get("name"));
                        sheetA.addCell(labelA);

                        System.out.println(""+(String) map.get("name")+" "+(String) map.get("phone"));
                    }
                    workbookA.write();    //写入数据        
                    workbookA.close();  //关闭连接
                    System.out.println("成功写入文件，请前往E盘查看文件！");
 
                } catch (Exception e) {
                    System.out.println("文件写入失败，报异常...");
                }
    }

    public static String uploadFile(String host,String fileName) {
        String line = null;
        String res = "";
        try {

            // 换行符
            final String newLine = "\r\n";
            final String boundaryPrefix = "--";
            // 定义数据分隔线
            String BOUNDARY = "========7d4a6d158c9";
            // 服务器的域名
            URL url = new URL(host);
            HttpURLConnection conn = (HttpURLConnection) url.openConnection();
            // 设置为POST情
            conn.setRequestMethod("POST");
            // 发送POST请求必须设置如下两行
            conn.setDoOutput(true);
            conn.setDoInput(true);
            conn.setUseCaches(false);
            // 设置请求头参数
            conn.setRequestProperty("connection", "Keep-Alive");
            conn.setRequestProperty("Charsert", "UTF-8");
            conn.setRequestProperty("Content-Type", "multipart/form-data; boundary=" + BOUNDARY);

            OutputStream out = new DataOutputStream(conn.getOutputStream());

            // 上传文件
            File file = new File(fileName);
            StringBuilder sb = new StringBuilder();
            sb.append(boundaryPrefix);
            sb.append(BOUNDARY);
            sb.append(newLine);
            // 文件参数,photo参数名可以随意修改
            sb.append("Content-Disposition: form-data;name=\"file\";filename=\"" + fileName
                    + "\"" + newLine);
            sb.append("Content-Type:application/octet-stream");
            // 参数头设置完以后需要两个换行，然后才是参数内容
            sb.append(newLine);
            sb.append(newLine);

            // 将参数头的数据写入到输出流中
            out.write(sb.toString().getBytes());

            // 数据输入流,用于读取文件数据
            DataInputStream in = new DataInputStream(new FileInputStream(
                    file));
            byte[] bufferOut = new byte[1024];
            int bytes = 0;
            // 每次读1KB数据,并且将文件数据写入到输出流中
            while ((bytes = in.read(bufferOut)) != -1) {
                out.write(bufferOut, 0, bytes);
            }
            // 最后添加换行
            out.write(newLine.getBytes());
            in.close();

            // 定义最后数据分隔线，即--加上BOUNDARY再加上--。
            byte[] end_data = (newLine + boundaryPrefix + BOUNDARY + boundaryPrefix + newLine)
                    .getBytes();
            // 写上结尾标识
            out.write(end_data);
            out.flush();
            out.close();

            // 定义BufferedReader输入流来读取URL的响应
            BufferedReader reader = new BufferedReader(new InputStreamReader(
                    conn.getInputStream()));

            while ((line = reader.readLine()) != null) {
                System.out.println(line);
                res += line;
            }


        } catch (Exception e) {
            System.out.println("发送POST请求出现异常！" + e);
            e.printStackTrace();
        }

        JSONObject json = JSONObject.parseObject(line);
        int code = json.getIntValue("code");
        if (code != 0){
            System.out.println("error : add person failed "+fileName);
            return "";
        }
        JSONObject data = json.getJSONObject("data");
        return data.getString("url");

    }


    /**
     * 发送HttpPost请求
     *
     * @param strURL
     *            服务地址
     * @param params
     *            json字符串,例如: "{ \"id\":\"12345\" }" ;其中属性名必须带双引号<br/>
     * @return 成功:返回json字符串<br/>
     */
    public static String post(String strURL, String params) {
        System.out.println(strURL);
        System.out.println(params);
        BufferedReader reader = null;
        try {
            URL url = new URL(strURL);// 创建连接
            HttpURLConnection connection = (HttpURLConnection) url.openConnection();
            connection.setDoOutput(true);
            connection.setDoInput(true);
            connection.setUseCaches(false);
            connection.setInstanceFollowRedirects(true);
            connection.setRequestMethod("POST"); // 设置请求方式
            // connection.setRequestProperty("Accept", "application/json"); // 设置接收数据的格式
            connection.setRequestProperty("Content-Type", "application/json"); // 设置发送数据的格式
            connection.connect();
            //一定要用BufferedReader 来接收响应， 使用字节来接收响应的方法是接收不到内容的
            OutputStreamWriter out = new OutputStreamWriter(connection.getOutputStream(), "UTF-8"); // utf-8编码
            out.append(params);
            out.flush();
            out.close();
            // 读取响应
            reader = new BufferedReader(new InputStreamReader(connection.getInputStream(), "UTF-8"));
            String line;
            String res = "";
            while ((line = reader.readLine()) != null) {
                res += line;
            }
            System.out.println("res:"+res);
            reader.close();


            //如果一定要使用如下方式接收响应数据， 则响应必须为: response.getWriter().print(StringUtils.join("{\"errCode\":\"1\",\"errMsg\":\"", message, "\"}")); 来返回
//            int length = (int) connection.getContentLength();// 获取长度
//            if (length != -1) {
//                byte[] data = new byte[length];
//                byte[] temp = new byte[512];
//                int readLen = 0;
//                int destPos = 0;
//                while ((readLen = is.read(temp)) > 0) {
//                    System.arraycopy(temp, 0, data, destPos, readLen);
//                    destPos += readLen;
//                }
//                String result = new String(data, "UTF-8"); // utf-8编码
//                System.out.println(result);
////                return result;
//            }

            return res;
        } catch (IOException e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        }
        return "error"; // 自定义错误信息
    }

}


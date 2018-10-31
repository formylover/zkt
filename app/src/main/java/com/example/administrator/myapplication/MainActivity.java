package com.example.administrator.myapplication;
import android.Manifest;
import android.content.Context;
import android.content.pm.PackageManager;
import android.graphics.Color;
import android.os.Build;
import android.os.Bundle;
import android.support.v4.app.ActivityCompat;
import android.support.v4.content.ContextCompat;
import android.support.v7.app.AppCompatActivity;
import android.text.SpannableString;
import android.text.Spanned;
import android.text.TextUtils;
import android.text.style.ForegroundColorSpan;
import android.util.Log;
import android.view.View;
import android.view.View.OnClickListener;
import android.view.inputmethod.InputMethodManager;
import android.widget.Button;
import android.widget.EditText;
import android.widget.TextView;
import java.text.SimpleDateFormat;

import java.io.File;
import java.util.Date;

import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.Workbook;

import android.os.Environment;

import android.widget.Toast;

public class MainActivity extends AppCompatActivity implements OnClickListener {
    private Button bt_1, bt_2;
    private String Str_inputtime;
    private int Int_inputtime, Int_Pass1;
    /**
     * 计算密码
     */
    private Button mButton;
    /**
     * 重置
     */
    private Button mButton2;
    /**  */

    private Button chaBtn;

    private TextView mTextView;
    private EditText mEditText;
    private EditText mEt;

    private File excelFile;
    private WritableWorkbook wwb;

    private String excelPath;

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_main);
        initView();

        //excelPath = getExcelDir()+File.separator+"demo.xls";
        //excelPath = "/stroage/emulated/0/test.xls";
        //excelFile = new File(excelPath);
        //createExcel(excelFile);

        //craetExcel();

        setTitle(R.string.MTitle);

        bt_1 = (Button) findViewById(R.id.button);
        bt_2 = (Button) findViewById(R.id.button2);

    }

    private void initView() {
        mButton = (Button) findViewById(R.id.button);
        mButton.setOnClickListener(this);
        mButton2 = (Button) findViewById(R.id.button2);
        mButton2.setOnClickListener(this);
        mEditText = (EditText) findViewById(R.id.editText);
        mEditText.setOnClickListener(this);
        mEditText.setEnabled(false);
        mEditText.setFocusable(false);
        mEt = (EditText) findViewById(R.id.et);
        mEt.setOnClickListener(this);
        chaBtn = (Button) findViewById(R.id.chaBtn);
        chaBtn.setOnClickListener(this);
    }

    @Override
    public void onClick(View v) {
        switch (v.getId()) {
            default:
                break;
            case R.id.button:

                //mEt.setText(getString(R.string.testclickbt));

                String Mst=mEt.getText().toString();
                if (TextUtils.isEmpty(Mst) )
                {
                    mEt.setFocusable(true);
                    InputMethodManager imm = (InputMethodManager)MainActivity.this.getSystemService(Context.INPUT_METHOD_SERVICE);
                    imm.toggleSoftInput(0, InputMethodManager.HIDE_NOT_ALWAYS);
                }
                else
                {
                    Int_inputtime=Integer.parseInt(Mst);
                    Int_Pass1=(9999-Int_inputtime)*(9999-Int_inputtime);

                    int Int_inputtime2,Int_inputtime3,Int_inputtime4,Int_inputtime5,Int_inputtime6,Int_inputtime7,Int_inputtime8;
                    int Int_pass2,Int_pass3,Int_pass4,Int_pass5,Int_pass6,Int_pass7,Int_pass8;
                    Int_inputtime2=Int_inputtime+1;
                    Int_inputtime3=Int_inputtime2+1;
                    Int_inputtime4=Int_inputtime3+1;
                    Int_inputtime5=Int_inputtime4+1;
                    Int_inputtime6=Int_inputtime5+1;
                    Int_inputtime7=Int_inputtime6+1;
                    Int_inputtime8=Int_inputtime7+1;

                    Int_pass2=(9999-Int_inputtime2)*(9999-Int_inputtime2);
                    Int_pass3=(9999-Int_inputtime3)*(9999-Int_inputtime3);
                    Int_pass4=(9999-Int_inputtime4)*(9999-Int_inputtime4);
                    Int_pass5=(9999-Int_inputtime5)*(9999-Int_inputtime5);
                    Int_pass6=(9999-Int_inputtime6)*(9999-Int_inputtime6);
                    Int_pass7=(9999-Int_inputtime7)*(9999-Int_inputtime7);
                    Int_pass8=(9999-Int_inputtime8)*(9999-Int_inputtime8);

                    //文本内容
                    String Sprin="1: "+Mst+" : "+String.valueOf(Int_Pass1)+"\n";
                    int strLen=Sprin.length();
                    SpannableString ss = new SpannableString("1: "+Mst+" : "+String.valueOf(Int_Pass1)+"\n"+"2: "+Int_inputtime2+" : "+String.valueOf(Int_pass2)+"\n"+"3: "+Int_inputtime3+" : "+String.valueOf(Int_pass3)+"\n"+"4: "+Int_inputtime4+" : "+String.valueOf(Int_pass4)+"\n"+"5: "+Int_inputtime5+" : "+String.valueOf(Int_pass5)+"\n"+"6: "+Int_inputtime6+" : "+String.valueOf(Int_pass6)+"\n"+"7: "+Int_inputtime7+" : "+String.valueOf(Int_pass7)+"\n"+"8: "+Int_inputtime8+" : "+String.valueOf(Int_pass8)+"\n");

                    //设置0-2的字符颜色
                    ss.setSpan(new ForegroundColorSpan(Color.RED), 0, strLen,Spanned.SPAN_EXCLUSIVE_EXCLUSIVE);

                    ss.setSpan(new ForegroundColorSpan(Color.BLUE), strLen, strLen*2,Spanned.SPAN_EXCLUSIVE_EXCLUSIVE);

                    ss.setSpan(new ForegroundColorSpan(Color.BLACK), strLen*2, strLen*3,Spanned.SPAN_EXCLUSIVE_EXCLUSIVE);

                    ss.setSpan(new ForegroundColorSpan(Color.GREEN), strLen*3, strLen*4,Spanned.SPAN_EXCLUSIVE_EXCLUSIVE);

                    ss.setSpan(new ForegroundColorSpan(Color.MAGENTA), strLen*4, strLen*5,Spanned.SPAN_EXCLUSIVE_EXCLUSIVE);

                    ss.setSpan(new ForegroundColorSpan(Color.CYAN), strLen*5, strLen*6,Spanned.SPAN_EXCLUSIVE_EXCLUSIVE);

                    ss.setSpan(new ForegroundColorSpan(Color.LTGRAY), strLen*6, strLen*7,Spanned.SPAN_EXCLUSIVE_EXCLUSIVE);

                    ss.setSpan(new ForegroundColorSpan(Color.DKGRAY), strLen*7, strLen*8,Spanned.SPAN_EXCLUSIVE_EXCLUSIVE);


                    //设置2-5的字符链接到电话簿，点击时触发拨号
                    //ss.setSpan(new URLSpan("tel:4155551212"), 18, 34,Spanned.SPAN_EXCLUSIVE_EXCLUSIVE);

                    //设置9-11的字符为网络链接，点击时打开页面
                    //ss.setSpan(new URLSpan("http://www.hao123.com"), 9, 11,Spanned.SPAN_EXCLUSIVE_EXCLUSIVE);

                    //设置13-15的字符点击时，转到写短信的界面，发送对象为10086
                    // ss.setSpan(new URLSpan("sms:10086"), 13, 15,Spanned.SPAN_EXCLUSIVE_EXCLUSIVE);

                    //粗体
                    //ss.setSpan(new StyleSpan(Typeface.BOLD_ITALIC), 5, 7,Spanned.SPAN_EXCLUSIVE_EXCLUSIVE);

                    //斜体
                    //ss.setSpan(new StrikethroughSpan(), 7, 10,Spanned.SPAN_EXCLUSIVE_EXCLUSIVE);

                    //mEditText.setText("1: "+Mst+" : "+String.valueOf(Int_Pass1)+"\n"+"2: "+Int_inputtime2+" : "+String.valueOf(Int_pass2)+"\n"+"3: "+Int_inputtime3+" : "+String.valueOf(Int_pass3)+"\n"+"4: "+Int_inputtime4+" : "+String.valueOf(Int_pass4)+"\n"+"5: "+Int_inputtime5+" : "+String.valueOf(Int_pass5)+"\n"+"6: "+Int_inputtime6+" : "+String.valueOf(Int_pass6)+"\n"+"7: "+Int_inputtime7+" : "+String.valueOf(Int_pass7)+"\n"+"8: "+Int_inputtime8+" : "+String.valueOf(Int_pass8)+"\n");
                    mEditText.setText(ss);
                }

                break;
            case R.id.button2:
                mEditText.setText("");
                mEt.setText("");
                mEt.setFocusable(true);
                mEt.setFocusableInTouchMode(true);
                mEt.requestFocus();
                InputMethodManager imm = (InputMethodManager)MainActivity.this.getSystemService(Context.INPUT_METHOD_SERVICE);
                imm.toggleSoftInput(0, InputMethodManager.HIDE_NOT_ALWAYS);
                break;
            case R.id.editText:
                break;
            case R.id.et:
                break;
            case R.id.chaBtn:
                excelPath = getExcelDir()+File.separator+"demo.xls";
                //excelPath = "/mnt/sdcard/test.xls";
                excelFile = new File(excelPath);

                if (Build.VERSION.SDK_INT>=Build.VERSION_CODES.M){
                    if (ContextCompat.checkSelfPermission(MainActivity.this,Manifest.permission.WRITE_EXTERNAL_STORAGE)!= PackageManager.PERMISSION_GRANTED){
                        //没有权限则申请权限
                        ActivityCompat.requestPermissions(MainActivity.this,new String[]{Manifest.permission.WRITE_EXTERNAL_STORAGE},1);
                    }else {
                        //有权限直接执行,docode()不用做处理
                        //doCode();
                        createExcel(excelFile);

                    }
                }else {
                    //小于6.0，不用申请权限，直接执行
                    //doCode();
                    createExcel(excelFile);
                }

                //createExcel(excelFile);

                //读取excel文件
                //readExcel();
                //更新excel文件
                //SimpleDateFormat   formatter   =   new   SimpleDateFormat   ("yyyy年MM月dd日 HH:mm:ss ");
                //Date curDate =  new Date(System.currentTimeMillis());
                //获取当前时间
                //String   str   =   formatter.format(curDate);

                //int getExcelrows=GetExcelRows();

                //1.) String s = String.valueOf(i);
                //
                //2.) String s = Integer.toString(i);
                //
                //3.) String s = "" + i;
                //Toast.makeText(this, "保存成功", Toast.LENGTH_SHORT).show();

                //writeToExcel(String.valueOf(getExcelrows),str);
                writeToExcel();
                //updatExcel();
                Toast.makeText(this, "当前时间保存成功", Toast.LENGTH_SHORT).show();
                break;
        }
    }

    /*
    private void craetExcel() {
        try {
            // 打开文件
            WritableWorkbook book = Workbook.createWorkbook(new File("mnt/sdcard/test.xls"));
            // 生成名为“第一张工作表”的工作表，参数0表示这是第一页
            WritableSheet sheet = book.createSheet("第一张工作表", 0);

            //设置行高,设置第一行高度为100,参数1：行数，参数2：高度
            sheet.setRowView(0, 100);
            //设置列宽,设置第一列宽度为50，参数1：列数，参数2：宽度
            sheet.setColumnView(0, 50);

            //合并单元格,参数1：合并的起始列数，参数2：合并的起始行数，参数3:合并的截止列数，参数4:合并的截止行数
            //合并第一列第一行到第四列第五行
            sheet.mergeCells(0, 0, 3, 4);

            //创建字体，参数1：字体样式，参数2：字号，参数3：粗体
            WritableFont font = new WritableFont(WritableFont.createFont("楷体"), 11, WritableFont.BOLD);
            WritableCellFormat format = new WritableCellFormat(font);

            //设置对齐方式为水平居中
            format.setAlignment(Alignment.CENTRE);
            //设置对齐方式为垂直居中
            format.setVerticalAlignment(VerticalAlignment.CENTRE);

            // 在Label对象的构造子中指名单元格位置是第一列第一行(0,0)
            // 以及单元格内容为baby
            Label label = new Label(0, 0, "baby", format);

            // 将定义好的单元格添加到工作表中
            sheet.addCell(label);
            // 生成一个保存数字的单元格，必须使用Number的完整包路径，否则有语法歧义。
            //单元格位置是第二列，第一行，值为123.456
            jxl.write.Number number = new jxl.write.Number(1, 0, 123.456);
            sheet.addCell(number);
            //写入数据并关闭
            book.write();
            book.close();

        } catch (WriteException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private void readExcel() {
        try {
            Workbook workbook = Workbook.getWorkbook(new File("mnt/sdcard/test.xls"));
            //获取第一个工作表的对象
            Sheet sheet = workbook.getSheet(0);
            //获取第一列第一行的的单元格
            Cell cell = sheet.getCell(0, 0);
            //获取单元格中的内容
            String body = cell.getContents();
            Log.d("bb", body);
            //读取数据关闭
            workbook.close();


        } catch (IOException e) {
            e.printStackTrace();
        } catch (BiffException e) {
            e.printStackTrace();
        }
    }

    private void updatExcel() {
        try {
            //获取excel文件
            Workbook workbook = Workbook.getWorkbook(new File("mnt/sdcard/test.xls"));
            //打开文件的副本，并且制定数据写回到原文件
            WritableWorkbook writableWorkbook = Workbook.createWorkbook(new File("mnt/sdcard/test.xls"), workbook);

            //修改工作表的数据
            WritableSheet sheet = writableWorkbook.getSheet(0);
            sheet.addCell(new Label(0, 0, "覆盖单元格内容"));

            //添加一张新的工作表
            WritableSheet sheet1 = writableWorkbook.createSheet("第二张工作表", 1);
            sheet1.addCell(new Label(0, 0, "第二页的数据"));

            //关闭
            writableWorkbook.close();
            workbook.close();

        } catch (IOException e) {
            e.printStackTrace();
        } catch (BiffException e) {
            e.printStackTrace();
        } catch (RowsExceededException e) {
            e.printStackTrace();
        } catch (WriteException e) {
            e.printStackTrace();
        }
    }
*/

    public void createExcel(File file) {
        WritableSheet ws = null;
        try {
            if (!file.exists()) {
                // 创建表
                wwb = Workbook.createWorkbook(file);
                // 创建表单,其中sheet表示该表格的名字,0表示第一个表格,
                ws = wwb.createSheet("sheet1", 0);
                // 在指定单元格插入数据
                Label lbl1 = new Label(0, 0, "序号");// 第一个参数表示,0表示第一列,第二个参数表示行,同样0表示第一行,第三个参数表示想要添加到单元格里的数据.
                Label bll2 = new Label(1, 0, "打卡记录时间");
                // 添加到指定表格里.
                ws.addCell(lbl1);
                ws.addCell(bll2);
                // 从内存中写入文件中
                wwb.write();
                wwb.close();
                //Toast.makeText(this, "打开文件成功！", Toast.LENGTH_SHORT).show();
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public void writeToExcel(String name, String gender) {
        try {
            //每次插入数据,都要取原来的表,然后新建一个表,然后将原来的表的内容添加到新表上.但只要两个路径相同的话,效果相当于在原来的表添加.
            Workbook oldWwb = Workbook.getWorkbook(excelFile);
            wwb = Workbook.createWorkbook(excelFile, oldWwb);
            //获取指定索引的表格
            WritableSheet ws = wwb.getSheet(0);
            // 获取该表格现有的行数
            int row = ws.getRows();
            Label lbl1 = new Label(0, row, name);
            Label bll2 = new Label(1, row, gender);
            ws.addCell(lbl1);
            ws.addCell(bll2);
            // 从内存中写入文件中,只能刷一次.
            wwb.write();
            wwb.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public void writeToExcel() {
        try {
            //每次插入数据,都要取原来的表,然后新建一个表,然后将原来的表的内容添加到新表上.但只要两个路径相同的话,效果相当于在原来的表添加.
            Workbook oldWwb = Workbook.getWorkbook(excelFile);
            wwb = Workbook.createWorkbook(excelFile, oldWwb);
            //获取指定索引的表格
            WritableSheet ws = wwb.getSheet(0);
            // 获取该表格现有的行数
            int row = ws.getRows();
            String Strrow=String.valueOf(row);

            SimpleDateFormat   formatter   =   new   SimpleDateFormat   ("yyyy年MM月dd日 HH:mm:ss ");
            Date curDate =  new Date(System.currentTimeMillis());
            //获取当前时间
            String   StrDate   =   formatter.format(curDate);

            Label lbl1 = new Label(0, row, Strrow);

            Label bll2 = new Label(1, row, StrDate);
            ws.addCell(lbl1);
            ws.addCell(bll2);
            // 从内存中写入文件中,只能刷一次.
            wwb.write();
            wwb.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public int GetExcelRows() {

        int getrows=0;
        try {
            //每次插入数据,都要取原来的表,然后新建一个表,然后将原来的表的内容添加到新表上.但只要两个路径相同的话,效果相当于在原来的表添加.
            Workbook oldWwb = Workbook.getWorkbook(excelFile);
            wwb = Workbook.createWorkbook(excelFile, oldWwb);
            //获取指定索引的表格
            WritableSheet ws = wwb.getSheet(0);
            // 获取该表格现有的行数

            getrows = ws.getRows();

            wwb.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
        return  getrows;
    }


    // 获取Excel文件夹
    public String getExcelDir() {
        // SD卡指定文件夹
        String sdcardPath = Environment.getExternalStorageDirectory()
                .toString();
        File dir = new File(sdcardPath + File.separator + File.separator + "Person");

        if (dir.exists()) {
            return dir.toString();

        } else {
            dir.mkdirs();
            Log.d("BAG", "保存路径不存在,");
            return dir.toString();
        }
    }

}
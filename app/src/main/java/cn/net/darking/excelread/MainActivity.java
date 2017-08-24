package cn.net.darking.excelread;

import android.content.Intent;
import android.net.Uri;
import android.os.Bundle;
import android.support.v7.app.AppCompatActivity;
import android.util.Log;
import android.view.View;
import android.widget.TextView;

import cn.net.darking.excelread.utils.POIExcelHelper;
import cn.net.darking.excelread.utils.PhotoOrVideoUtils;

import static cn.net.darking.excelread.utils.PhotoOrVideoUtils.REQUEST_CODE_GET_FILES;

public class MainActivity extends AppCompatActivity implements View.OnClickListener {

    private TextView text;

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_main);

        text = (TextView) findViewById(R.id.text);

        text.setOnClickListener(this);
    }

    @Override
    public void onClick(View v) {
        Intent intent = new Intent();
        //intent.setType("*/*");
        intent.setType("*/*");
        //intent.setType("text/plain;application/msword;application/vnd.ms-powerpoint;image/*;video/*");
        intent.addCategory(Intent.CATEGORY_OPENABLE);
        intent.setAction(Intent.ACTION_GET_CONTENT);
        startActivityForResult(intent, REQUEST_CODE_GET_FILES);
    }


    @Override
    protected void onActivityResult(int requestCode, int resultCode, Intent data) {
        super.onActivityResult(requestCode, resultCode, data);
        if (data != null) {
            Uri uri = PhotoOrVideoUtils.getFileUri(requestCode, resultCode, data);
            String path = PhotoOrVideoUtils.getPath(this, uri);

            Log.e("文件路径",path);

//            ExcelHelper helper = new ExcelHelper();
//
//            try {
//                FileInputStream fis = new FileInputStream(path);
//                String[] titles = helper.readExcelTitle(fis);
//
//                for (String title : titles) {
//                    Log.e("文件标题",title);
//                }
//
//
//
//
//
//            } catch (FileNotFoundException e) {
//                e.printStackTrace();
//            }

            POIExcelHelper poiExcelHelper = new POIExcelHelper(path);





        }

    }
}

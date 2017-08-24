package cn.net.darking.excelread.utils;

import android.annotation.TargetApi;
import android.content.ContentUris;
import android.content.Context;
import android.content.Intent;
import android.database.Cursor;
import android.graphics.Bitmap;
import android.net.Uri;
import android.os.Build;
import android.os.Bundle;
import android.os.Environment;
import android.provider.DocumentsContract;
import android.provider.MediaStore;
import android.support.v4.app.FragmentActivity;
import android.text.format.DateFormat;
import android.util.Log;
import android.widget.PopupWindow;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Calendar;
import java.util.Locale;

import static android.app.Activity.RESULT_CANCELED;
import static android.app.Activity.RESULT_OK;

/**
 * 获取照片的工具类
 * Created by 靳宁 on 2017/2/14.
 */

public class PhotoOrVideoUtils {
    public static int REQUEST_CODE_TAKE_PHOTO = 100;//拍照
    public static int REQUEST_CODE_GET_PHOTO = 200;//从相片中获取照片
    public static int REQUEST_CODE_TAKE_VIDEO = 300;//拍摄视频
    public static int REQUEST_CODE_GET_VIDEOORPHOTO = 400;//从相册中获取照片或者视频
    public static int REQUEST_CODE_GET_FILES = 500;
    private static String BASE_PATH = "";

    private static PopupWindow photoWindow, videoWindow;

    public static FragmentActivity getActivity() {
        return activity;
    }

    public static void setActivity(FragmentActivity activity) {
        PhotoOrVideoUtils.activity = activity;
    }


    private static FragmentActivity activity;


    public static void doFiles(FragmentActivity activity) {


//        Intent intent = new Intent(Intent.ACTION_GET_CONTENT);
//        intent.setType("*/*");
//        intent.addCategory(Intent.CATEGORY_OPENABLE);
        setActivity(activity);
        Intent intent = new Intent();
        //intent.setType("*/*");
        intent.setType("*/*");
        //intent.setType("text/plain;application/msword;application/vnd.ms-powerpoint;image/*;video/*");
        intent.addCategory(Intent.CATEGORY_OPENABLE);
        intent.setAction(Intent.ACTION_GET_CONTENT);
        getActivity().startActivityForResult(intent, REQUEST_CODE_GET_FILES);
    }


    public static Uri getFileUri(int requestCode, int resultCode, Intent data) {
        Uri uri = null;
        //获取到了图片或者视频，并设置图片

        if (requestCode == REQUEST_CODE_GET_VIDEOORPHOTO) {
            if (data != null) {
//                String path = "";
                if (getAndroidOSVersion() < 19) {
                    if (data != null) {
                        uri = data.getData();
//                        path = getPath(getActivity(), uri);
                    }
                } else {
                    if (resultCode == RESULT_OK && data != null) {
                        uri = data.getData();
//                        path = getPath(getActivity(), uri);
                        //LogUtil.e("视频路径为" + getPath(getActivity(), uri));
                    }
                }
//                if (path.endsWith(".jpg") || path.endsWith(".png") || path.endsWith(".bmp")) {
//
//                } else if (path.endsWith(".rmvb") || path.endsWith(".3gp") || path.endsWith(".mp4") || path.endsWith(".mov")) {
//
//                } else {
//                    //Toast.makeText(getActivity(), "请选择正确的图像或者视频模式！", Toast.LENGTH_LONG).show();
//                }
            }

        } else if (requestCode == REQUEST_CODE_TAKE_PHOTO) {

            if (data != null) {
                if (data.getExtras() != null /*|| data.getExtras().get("data").equals("")*/) {
                    if (data.getExtras().get("data") != null) {

                        Log.e("拍摄图片路径为", data.getExtras().get("data").toString());
                        //  相机照相获取文件路径开始
                        String sdStatus = Environment.getExternalStorageState();
                        getSavePath();
                        Bitmap bitmap = null;
                        String fileName = "";
                        Bundle bundle = data.getExtras();
                        String name = DateFormat.format("yyyyMMdd_hhmmss", Calendar.getInstance(Locale.CHINA)) + ".jpg";
                        //Toast.makeText(getContext(), name, Toast.LENGTH_LONG).show();
                        new DateFormat();

                        bitmap = (Bitmap) bundle.get("data");// 获取相机返回的数据，并转换为Bitmap图片格式
                        //imageView.setImageBitmap(bitmap);// 将图片显示在ImageView里
                        FileOutputStream b = null;
                        fileName = getSavePath() + name;
                        File file = new File(getSavePath());
                        boolean m = file.mkdirs();// 创建文件夹


                        try {
                            Log.e("保存文件地址", fileName);
                            b = new FileOutputStream(fileName);//这里应该是文件名,而不是一个文件夹的名称
//                            LogUtil.e("保存文件地址-------+++++++------------"+b);
                            bitmap.compress(Bitmap.CompressFormat.JPEG, 100, b);// 把数据写入文件--参数100，即压缩品质为100%
                            //TODO 获取文件的URI  Uri uri = Uri.parse("file:///sdcard/img.png");
                            uri = Uri.fromFile(new File(fileName));

                            //uri = data.getData();
                        } catch (FileNotFoundException e) {
                            e.printStackTrace();
                        } finally {
                            bitmap.recycle();//释放掉bitmap
                            try {
                                b.flush();
                                b.close();
                            } catch (IOException e) {
                                e.printStackTrace();
                            }
                        }
                        Log.e("", "拍摄了照片");

                    }
                }
            }
            //  相机照相获取文件路径结束
        } else if (requestCode == REQUEST_CODE_TAKE_VIDEO) {
            Log.e("", "拍摄了视频");
            if (resultCode == RESULT_OK) {
                // Video captured and saved to fileUri specified in the Intent
                Log.e("拍摄视频的路径", data.getData().getPath());
                //Display the video
                uri = data.getData();
                Log.e("视频路径为", getPath(getActivity(), uri));
                // setFilePath(mPath);
            } else if (resultCode == RESULT_CANCELED) {
                // User cancelled the video capture
                uri = null;
            }
        } else if (requestCode == REQUEST_CODE_GET_PHOTO) {
            if (data != null) {//当进入选择图片界面，没有选择图片而退出的时候，data返回值会为null
                uri = data.getData();
            }
        } else if (requestCode == REQUEST_CODE_GET_FILES) {
            if (data != null) {//当进入选择文件界面，没有选择图片而退出的时候，data返回值会为null
                uri = data.getData();
            }
        }
        return uri;
    }


    /**
     * 获取相册中的图片、视频绝对路径
     *
     * @param context getContext().getApplicationContext()
     * @param uri     Intent data.getData()
     * @return 相册中的图片、视频绝对路径
     */
    @TargetApi(Build.VERSION_CODES.KITKAT)
    public static String getPath(final Context context, final Uri uri) {
        final boolean isKitKat = Build.VERSION.SDK_INT >= Build.VERSION_CODES.KITKAT;
        // DocumentProvider
        if (isKitKat && DocumentsContract.isDocumentUri(context, uri)) {
            // ExternalStorageProvider
            if (isExternalStorageDocument(uri)) {
                final String docId = DocumentsContract.getDocumentId(uri);
                final String[] split = docId.split(":");
                final String type = split[0];
                if ("primary".equalsIgnoreCase(type)) {
                    return Environment.getExternalStorageDirectory() + "/"
                            + split[1];
                }
                // TODO handle non-primary volumes
            }
            // DownloadsProvider
            else if (isDownloadsDocument(uri)) {
                final String id = DocumentsContract.getDocumentId(uri);
                final Uri contentUri = ContentUris.withAppendedId(
                        Uri.parse("content://downloads/public_downloads"),
                        Long.valueOf(id));
                return getDataColumn(context, contentUri, null, null);
            }
            // MediaProvider
            else if (isMediaDocument(uri)) {
                final String docId = DocumentsContract.getDocumentId(uri);
                final String[] split = docId.split(":");
                final String type = split[0];
                Uri contentUri = null;
                if ("image".equals(type)) {
                    contentUri = MediaStore.Images.Media.EXTERNAL_CONTENT_URI;
                } else if ("video".equals(type)) {
                    contentUri = MediaStore.Video.Media.EXTERNAL_CONTENT_URI;
                } else if ("audio".equals(type)) {
                    contentUri = MediaStore.Audio.Media.EXTERNAL_CONTENT_URI;
                }
                final String selection = MediaStore.MediaColumns._ID + "=?";
                final String[] selectionArgs = new String[]{split[1]};
                return getDataColumn(context, contentUri, selection,
                        selectionArgs);
            }
        }
        // MediaStore (and general)
        else if ("content".equalsIgnoreCase(uri.getScheme())) {
            // Return the remote address
            if (isGooglePhotosUri(uri))
                return uri.getLastPathSegment();
            return getDataColumn(context, uri, null, null);
        }
        // File
        else if ("file".equalsIgnoreCase(uri.getScheme())) {
            return uri.getPath();
        }
        return null;
    }

    /**
     * Get the value of the data column for this Uri . This is useful for
     * MediaStore Uris , and other file - based ContentProviders.
     *
     * @param context       The context.
     * @param uri           The Uri to query.
     * @param selection     (Optional) Filter used in the query.
     * @param selectionArgs (Optional) Selection arguments used in the query.
     * @return The value of the _data column, which is typically a file path.
     */
    public static String getDataColumn(Context context, Uri uri,
                                       String selection, String[] selectionArgs) {
        Cursor cursor = null;
        final String column = MediaStore.MediaColumns.DATA;
        final String[] projection = {column};
        try {
            cursor = context.getContentResolver().query(uri, projection,
                    selection, selectionArgs, null);
            if (cursor != null && cursor.moveToFirst()) {
                final int index = cursor.getColumnIndexOrThrow(column);
                return cursor.getString(index);
            }
        } finally {
            if (cursor != null)
                cursor.close();
        }
        return null;
    }

    /**
     * @param uri The Uri to check.
     * @return Whether the Uri authority is ExternalStorageProvider.
     */
    public static boolean isExternalStorageDocument(Uri uri) {
        return "com.android.externalstorage.documents".equals(uri
                .getAuthority());
    }

    /**
     * @param uri The Uri to check.
     * @return Whether the Uri authority is DownloadsProvider.
     */
    public static boolean isDownloadsDocument(Uri uri) {
        return "com.android.providers.downloads.documents".equals(uri
                .getAuthority());
    }

    /**
     * @param uri The Uri to check.
     * @return Whether the Uri authority is MediaProvider.
     */
    public static boolean isMediaDocument(Uri uri) {
        return "com.android.providers.media.documents".equals(uri
                .getAuthority());
    }

    /**
     * @param uri The Uri to check.
     * @return Whether the Uri authority is Google Photos.
     */
    public static boolean isGooglePhotosUri(Uri uri) {
        return "com.google.android.apps.photos.content".equals(uri
                .getAuthority());
    }

    /**
     * 获取SDK版本号
     *
     * @return
     */
    public static int getAndroidOSVersion() {
        int osVersion;
        try {
            osVersion = Integer.valueOf(Build.VERSION.SDK);
        } catch (NumberFormatException e) {
            osVersion = 0;
        }

        return osVersion;
    }

    /**
     * 指定保存照片或者视频的文件夹（绝对路径）
     *
     * @return 绝对路径地址
     */
    public static String getSavePath() {
        String path = "";
        if (Environment.getExternalStorageState().equals(Environment.MEDIA_MOUNTED)) {
            //一定要指定文件名，不然会将下载的文件保存成为.tmp的文件
            path = Environment.getExternalStorageDirectory().getPath() + File.separator + "shuxin" + File.separator + "picture" + File.separator;
        } else {
            path = getActivity().getFilesDir().getAbsolutePath() + File.separator;
            //Toast.makeText(getActivity(), "SD卡不可用,将保存到手机内部", Toast.LENGTH_SHORT).show();
        }
        return path;
    }

    /**
     * 文件类型列表
     */
    static String[][] typeStr = {
            {".wps", ".wpt", ".doc", ".dot", ".rtf", ".docx", ".dotx"}/*word类型0*/
            , {".et", ".ett", ".xls", ".xlt", ".xlsx", ".xlsm"}/*excl文件类型1*/
            , {".dps", ".dpt", ".pptx", ".ppt", ".pot", ".pps"}/*ppt文件类型2*/
            , {".bmp", ".jpg", ".jpeg", ".tif", ".png", ".pcx", ".ico", ".cur"}/*图片文件类型3*/
            , {".avi", ".rm", ".rmvb", ".mp4", ".3gp", ".mov", ".wmv"}/*视频文件类型4*/
            , {".cd", ".ogg", ".mp3", ".asf", ".wma", ".wav", ".mp3pro", ".ape"}/*音频文件类型5*/};

    /**
     * 获取文件类型
     *
     * @param path
     * @return 0：word类型--1：excl类型--2：ppt类型--3：图片类型--4：视频类型--5：音频类型
     * -1：未知类型
     */
    public static int getLastnameType(String path) {
        path = path.toLowerCase();
        Log.e("$$$path$$$", path);
        path = path.substring(path.lastIndexOf("/"));//不包含 (/)线
        Log.e("$$$path$$$", path);
        path = path.substring(path.lastIndexOf("."));//  不包含  (.)
        Log.e("$$$path$$$", path);
        int type = -1;
        for (int i = 0; i < 6; i++) {
            for (int j = 0; j < typeStr[i].length; j++) {
                Log.e("---typeStr---", typeStr[i][j] + "===" + path.equals(typeStr[i][j].toString()));

                if (path.equals(typeStr[i][j].toString())) {
                    return i;
                } else {
                    type = -1;
                }
            }
        }

        return type;
    }
}

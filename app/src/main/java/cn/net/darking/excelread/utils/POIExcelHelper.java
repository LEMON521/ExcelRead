package cn.net.darking.excelread.utils;

import android.util.Log;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

/**
 * Created by Zrzc on 2017/8/24.
 */

public class POIExcelHelper {
    private FileInputStream fis = null;
    private int fileType = -1;
    private final int OTHER = -1;
    private final int Excel = 1;
    private final int Word = 2;
    private final int PowerPoint = 3;
    private final int Visio = 4;

    public POIExcelHelper(String path) {


        try {
            fis = new FileInputStream(path);
            POIFSFileSystem fs = new POIFSFileSystem(fis);
            //得到Excel工作簿对象
            HSSFWorkbook wb = new HSSFWorkbook(fs);
            for (int i = 0; i < wb.getNumberOfSheets(); i++) {
                //得到Excel工作表对象
                HSSFSheet sheet = wb.getSheetAt(i);
                String sheetName = sheet.getSheetName();


                Log.e("表名", sheetName);
                Log.e("测试", "================");
                Log.e("第", sheet.getFirstRowNum() + "行");
                Log.e("共", sheet.getLastRowNum() + "行");
                Log.e("测试", "================");

                int rowNum = 0;
                while (sheet.rowIterator().hasNext()) {
                    if (rowNum == sheet.getLastRowNum() + 1) {
                        break;//如果为最后一行,则跳出--(最后一行指的是有数据的最后一行)
                    }
                    Log.e("行", rowNum + "");
                    //得到Excel工作表的行
                    HSSFRow row = sheet.getRow(rowNum++);

                    for (int cellNum = 0; cellNum < row.getLastCellNum(); cellNum++) {
                        HSSFCell cell = row.getCell((short) cellNum);
                        if (cell != null)
                            Log.e(cellNum + "列", cell.toString());
                    }
                    sheet.rowIterator().next();
                }
//
//                for (int rowNum = 0; rowNum < sheet.getLastRowNum(); rowNum++) {//一共多少行
//
//
//                }
            }


        } catch (FileNotFoundException e) {
            e.printStackTrace();
            Log.e("FileNotFoundException错误", e.toString());
        } catch (IOException e) {
            e.printStackTrace();
            Log.e("IOException错误", e.toString());
        }


    }


}

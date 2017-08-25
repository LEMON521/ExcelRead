package cn.net.darking.excelread.utils;

import android.util.Log;

import org.apache.poi.POIOLE2TextExtractor;
import org.apache.poi.POITextExtractor;
import org.apache.poi.extractor.ExtractorFactory;
import org.apache.poi.hdgf.extractor.VisioTextExtractor;
import org.apache.poi.hslf.extractor.PowerPointExtractor;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hwpf.extractor.WordExtractor;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.extractor.ExcelExtractor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.xmlbeans.XmlException;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.HashMap;

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


    public POIExcelHelper() {

    }


    public HashMap<String, HashMap<String, HashMap<String, HashMap<String, HashMap<String, String>>>>> getExcel2003Datas(String[] paths, String[] tableType) {


        //1:文件
        //2:根据文件类型来判断对应的哪张表
        //3:文件的工作簿
        //4:行
        //5:列

        //总的数据
        HashMap<String, HashMap<String, HashMap<String, HashMap<String, HashMap<String, String>>>>> totalMap = new HashMap<>();

        //table类型
        HashMap<String, HashMap<String, HashMap<String, HashMap<String, String>>>> tableMap = null;

        //sheet
        HashMap<String, HashMap<String, HashMap<String, String>>> sheetMap = null;

        //row
        HashMap<String, HashMap<String, String>> rowMap = null;

        //cell
        HashMap<String, String> cellMap = null;

        //每次用之前,清空数据
//        cellMap.clear();
//        rowMap.clear();
//        sheetMap.clear();
//        tableMap.clear();
//        totalMap.clear();

        for (int index = 0; index < paths.length; index++) {
            try {
                fis = new FileInputStream(paths[index]);
                POIFSFileSystem fs = new POIFSFileSystem(fis);
                //得到Excel工作簿对象
                HSSFWorkbook wb = new HSSFWorkbook(fs);
                tableMap = new HashMap<>();
                sheetMap = new HashMap<>();
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
                    rowMap = new HashMap<>();
                    while (sheet.rowIterator().hasNext()) {
                        if (rowNum == sheet.getLastRowNum() + 1) {
                            break;//如果为最后一行,则跳出--(最后一行指的是有数据的最后一行)
                        }
                        Log.e("行", rowNum + "");
                        //得到Excel工作表的行
                        HSSFRow row = sheet.getRow(rowNum);
                        cellMap = new HashMap<>();
                        if (row != null) {
                            for (int cellNum = 0; cellNum < row.getLastCellNum(); cellNum++) {
                                HSSFCell cell = row.getCell((short) cellNum);
                                if (cell != null) {
                                    cellMap.put(cellNum + "", cell.toString());
                                    Log.e(cellNum + "列", cell.toString());
                                }
                            }
                            if (cellMap.size() > 0) {//此行有数据,才添加
                                rowMap.put(rowNum + "", cellMap);
                            }
                        }
                        rowNum++;
                        sheet.rowIterator().next();
                    }
                    sheetMap.put(sheetName, rowMap);
                }
                tableMap.put(tableType[index], sheetMap);

                totalMap.put(paths[index], tableMap);


            } catch (FileNotFoundException e) {
                e.printStackTrace();
                Log.e("FileNotFoundException错误", e.toString());
                return null;

            } catch (IOException e) {
                e.printStackTrace();
                Log.e("IOException错误", e.toString());
                return null;
            } finally {
                try {
                    fis.close();//关闭流,释放资源
                    Log.e("关闭流", "成功关闭");
                } catch (IOException e) {
                    e.printStackTrace();
                    Log.e("关闭流", e.toString());
                }
            }
        }

        return totalMap;
    }


    public HashMap<String, HashMap<String, HashMap<String, HashMap<String, HashMap<String, String>>>>> getExcelDatas(String[] paths, String[] tableType) {
        //1:文件
        //2:根据文件类型来判断对应的哪张表
        //3:文件的工作簿
        //4:行
        //5:列

        //总的数据
        HashMap<String, HashMap<String, HashMap<String, HashMap<String, HashMap<String, String>>>>> totalMap = new HashMap<>();

        //table类型
        HashMap<String, HashMap<String, HashMap<String, HashMap<String, String>>>> tableMap = null;

        //sheet
        HashMap<String, HashMap<String, HashMap<String, String>>> sheetMap = null;

        //row
        HashMap<String, HashMap<String, String>> rowMap = null;

        //cell
        HashMap<String, String> cellMap = null;

        Workbook wb = null;


        for (int index = 0; index < paths.length; index++) {
            try {
                fis = new FileInputStream(paths[index]);
                if (paths[index].endsWith(".xls")) {
                    wb = new HSSFWorkbook(fis);
                } else if (paths[index].endsWith("xlsx")) {
                    wb = new XSSFWorkbook(fis);
                } else {
                    System.out.println("您输入的excel格式不正确");
                    return null;
                }

                //得到Excel工作簿对象
//                HSSFWorkbook wb = new HSSFWorkbook(fs);
                tableMap = new HashMap<>();
                sheetMap = new HashMap<>();
                for (int i = 0; i < wb.getNumberOfSheets(); i++) {
                    //得到Excel工作表对象
                    Sheet sheet = wb.getSheetAt(i);
                    String sheetName = sheet.getSheetName();


                    Log.e("表名", sheetName);
                    Log.e("测试", "================");
                    Log.e("第", sheet.getFirstRowNum() + "行");
                    Log.e("共", sheet.getLastRowNum() + "行");
                    Log.e("测试", "================");

                    int rowNum = 0;
                    rowMap = new HashMap<>();
                    while (sheet.rowIterator().hasNext()) {
                        if (rowNum == sheet.getLastRowNum() + 1) {
                            break;//如果为最后一行,则跳出--(最后一行指的是有数据的最后一行)
                        }
                        Log.e("行", rowNum + "");
                        //得到Excel工作表的行
                        Row row = sheet.getRow(rowNum);
                        cellMap = new HashMap<>();
                        if (row != null) {
                            for (int cellNum = 0; cellNum < row.getLastCellNum(); cellNum++) {
                                Cell cell = row.getCell((short) cellNum);
                                if (cell != null) {
                                    cellMap.put(cellNum + "", cell.toString());
                                    Log.e(cellNum + "列", cell.toString());
                                }
                            }
                            if (cellMap.size() > 0) {//此行有数据,才添加
                                rowMap.put(rowNum + "", cellMap);
                            }
                        }
                        rowNum++;
                        sheet.rowIterator().next();
                    }
                    sheetMap.put(sheetName, rowMap);
                }
                tableMap.put(tableType[index], sheetMap);

                totalMap.put(paths[index], tableMap);


            } catch (FileNotFoundException e) {
                e.printStackTrace();
                Log.e("FileNotFoundException错误", e.toString());
                return null;

            } catch (IOException e) {
                e.printStackTrace();
                Log.e("IOException错误", e.toString());
                return null;
            } finally {
                try {
                    fis.close();//关闭流,释放资源
                    Log.e("关闭流", "成功关闭");
                } catch (IOException e) {
                    e.printStackTrace();
                    Log.e("关闭流", e.toString());
                }
            }
        }

        return totalMap;


    }

    public void test(String path) throws IOException, OpenXML4JException, XmlException {

        FileInputStream fis = new FileInputStream(path);
        POIFSFileSystem fileSystem = new POIFSFileSystem(fis);
        // Firstly, get an extractor for the Workbook
        POIOLE2TextExtractor oleTextExtractor =
                ExtractorFactory.createExtractor(fileSystem);
        // Then a List of extractors for any embedded Excel, Word, PowerPoint
        // or Visio objects embedded into it.
        POITextExtractor[] embeddedExtractors =
                ExtractorFactory.getEmbededDocsTextExtractors(oleTextExtractor);
        for (POITextExtractor textExtractor : embeddedExtractors) {
            // If the embedded object was an Excel spreadsheet.
            if (textExtractor instanceof ExcelExtractor) {
                ExcelExtractor excelExtractor = (ExcelExtractor) textExtractor;
                Log.e("什么文件===Excel=", excelExtractor.getText());
                System.out.println(excelExtractor.getText());
            }
            // A Word Document
            else if (textExtractor instanceof WordExtractor) {
                WordExtractor wordExtractor = (WordExtractor) textExtractor;
                String[] paragraphText = wordExtractor.getParagraphText();
                for (String paragraph : paragraphText) {
                    Log.e("什么文件===Word=", paragraph);
                    System.out.println(paragraph);
                }
                // Display the document's header and footer text
                Log.e("Word=页脚", wordExtractor.getFooterText());
                Log.e("Word=页眉", wordExtractor.getHeaderText());
                System.out.println("Footer text: " + wordExtractor.getFooterText());
                System.out.println("Header text: " + wordExtractor.getHeaderText());
            }
            // PowerPoint Presentation.
            else if (textExtractor instanceof PowerPointExtractor) {
                PowerPointExtractor powerPointExtractor =
                        (PowerPointExtractor) textExtractor;
                Log.e("PowerPoint=Text", powerPointExtractor.getText());
                Log.e("PowerPoint=Notes", powerPointExtractor.getNotes());
                System.out.println("Text: " + powerPointExtractor.getText());
                System.out.println("Notes: " + powerPointExtractor.getNotes());
            }
            // Visio Drawing
            else if (textExtractor instanceof VisioTextExtractor) {
                VisioTextExtractor visioTextExtractor =
                        (VisioTextExtractor) textExtractor;
                Log.e("Visio=Text", visioTextExtractor.getText());
                System.out.println("Text: " + visioTextExtractor.getText());
            }
        }


    }
}

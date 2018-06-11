package com.wzw.excel;

import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.*;

/**
 * @Author:WangZhiwen
 * @Description:
 * @Date:2018-1-7
 */
public class SimpleExcel {

    private static Logger LOGGER = LoggerFactory.getLogger(SimpleExcel.class);

    private SXSSFWorkbook workbook = new SXSSFWorkbook();

    private XSSFCellStyle headerStyle;

    private XSSFCellStyle contentStyle;

    private Map<String,SXSSFSheet> sheetMap = new HashMap<String,SXSSFSheet>();//多个sheet



    public boolean addSheet(String sheetName){
        SXSSFSheet sheet = workbook.createSheet(sheetName);
        sheetMap.put(sheetName,sheet);
        return true;
    }

    public void writeTitle(List<String> title, String sheetName) {
        SXSSFSheet sheet = sheetMap.get(sheetName);
        if(sheet==null){
            LOGGER.error("指定名称的sheet不存在");
            return;
        }
        SXSSFRow row = sheet.createRow(0);
        fillRow(row,title);
    }

    /**
     * 向指定name的sheet中写入多行内容
     * @param lines
     * @param sheetName
     */
    public void writeContent(List<List<String>> lines, String sheetName) {
        SXSSFSheet sheet = sheetMap.get(sheetName);
        if(sheet==null){
            LOGGER.error("指定名称的sheet不存在");
            return;
        }
        int i = 0;
        Iterator lineIterator = lines.iterator();
        while(lineIterator.hasNext()) {
            ++i;
            List<String> line = (List<String>)lineIterator.next();
            SXSSFRow row = sheet.createRow(i);
            fillRow(row,line);
        }
    }

    /**
     * 向指定name的sheet中写入一行内容
     * @param line
     * @param sheetName
     */
    public void appendLine(List<String> line, String sheetName) {
        SXSSFSheet sheet = sheetMap.get(sheetName);

        if(sheet==null){
            LOGGER.error("指定名称的sheet不存在");
            return;
        }

        LOGGER.info("getLastFlushedRowNum:{},getLastRowNum:{}",sheet.getLastFlushedRowNum(),sheet.getLastRowNum());
        int lastRowNum = sheet.getLastRowNum();

        LOGGER.info("开始写入第{}行内容",lastRowNum+1);
        SXSSFRow row = sheet.createRow(lastRowNum+1);
        for(int i=0;i<line.size();i++) {
            SXSSFCell cell = row.createCell(i);
            cell.setCellValue(new XSSFRichTextString(line.get(i)));
            cell.setCellStyle(contentStyle);
        }
        LOGGER.info("成功写入第{}行内容",lastRowNum+1);
    }

    public void output(String path){
        File excelFile = new File(path);
        try {
            workbook.write(new FileOutputStream(excelFile));
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private void fillRow(SXSSFRow row, List<String> content){
        int i=0;
        Iterator cellIterator = content.iterator();
        while(cellIterator.hasNext()){
            String value = (String)cellIterator.next();
            SXSSFCell cell = row.createCell(i++);
            cell.setCellValue(new XSSFRichTextString(value));
            cell.setCellStyle(headerStyle);
        }
    }

    private XSSFCellStyle defaultHeaderStyle() {
        XSSFCellStyle style = (XSSFCellStyle)workbook.createCellStyle();
        style.setAlignment((short)2);
        style.setVerticalAlignment((short)1);
        style.setWrapText(true);
        XSSFFont font = (XSSFFont)workbook.createFont();
        font.setBoldweight((short)700);
        font.setFontName("宋体");
        font.setFontHeight(200);
        font.setFontHeightInPoints((short)11);
        style.setFont(font);
        return style;
    }

    private XSSFCellStyle defaultContentStyle() {
        XSSFCellStyle style = (XSSFCellStyle)workbook.createCellStyle();
        style.setVerticalAlignment((short)1);
        XSSFFont font = (XSSFFont)workbook.createFont();
        font.setFontName("宋体");
        style.setFont(font);
        return style;
    }

    public SimpleExcel() {
        /*try {
            this.workbook = new SXSSFWorkbook();
        } catch (Exception e) {
            e.printStackTrace();
        }
        int sheetCount = workbook.getNumberOfSheets();
        for(int i=0;i<sheetCount;i++){
            SXSSFSheet sheet = workbook.getSheetAt(i);
            sheetMap.put(sheet.getSheetName(),sheet);
        }*/
        this.headerStyle = defaultHeaderStyle();
        this.contentStyle = defaultContentStyle();
    }

    public static void main(String[] args) {

        SimpleExcel SimpleExcel = new SimpleExcel();
        SimpleExcel.addSheet("测试");

        List<String> line = new ArrayList<String>();
        line.add("测试");

        SimpleExcel.appendLine(line,"测试");
        SimpleExcel.appendLine(line,"测试");

    }


}

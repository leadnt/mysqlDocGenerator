package com.guitar.dbdoc;

import java.io.File;
import java.io.FileInputStream;
import java.util.List;
import java.util.Properties;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.util.CellRangeAddress;

/**
 *
 * @author hxy
 */
public class DocFactory {

    private final Properties prop;

    ExcelUtil excel;
    DataUtil dataUtil;
    Font titleFont = null;
    Font defaultFont = null;
    CellStyle titleStyle = null;
    CellStyle defaultStyle = null;

    public DocFactory(File cfg) throws Exception {
        prop = new Properties();
        prop.load(new FileInputStream(cfg));

    }

    public void generator() throws Exception {

        String url = prop.getProperty("url");
        String db = prop.getProperty("db");
        String user = prop.getProperty("user");
        String pwd = prop.getProperty("pwd");

        excel = new ExcelUtil();
        dataUtil = new DataUtil();
        //book = new HSSFWorkbook();
        titleFont = excel.createFont();
        titleFont.setFontName("宋体");
        titleFont.setFontHeightInPoints((short) 11);
        titleFont.setColor(HSSFColor.WHITE.index);

        defaultFont = excel.createFont();
        defaultFont.setFontName("宋体");
        defaultFont.setFontHeightInPoints((short) 11);

        titleStyle = excel.createStyle();
        titleStyle.setFillForegroundColor(HSSFColor.ROYAL_BLUE.index);
        titleStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);
        titleStyle.setFont(titleFont);
        excel.setStyleBorder(titleStyle, "1111", HSSFCellStyle.BORDER_THIN);
        excel.setStyleBorderColor(titleStyle, "1111", HSSFColor.PALE_BLUE.index);

        defaultStyle = excel.createStyle();
        defaultStyle.setFont(defaultFont);
        excel.setStyleBorder(defaultStyle, "1111", HSSFCellStyle.BORDER_THIN);
        excel.setStyleBorderColor(defaultStyle, "1111", HSSFColor.PALE_BLUE.index);

        HSSFSheet sheet = excel.getSheet("目录", true);
        sheet.setDefaultRowHeightInPoints(20);
        setColumnWidth(sheet, 0, 10);
        setColumnWidth(sheet, 1, 40);
        setColumnWidth(sheet, 2, 60);

        List<Table> tables = dataUtil.queryTables(url+db, user, pwd);
        int sheetrow = 0;
        excel.setCellValue(sheet, sheetrow, 0, "序号", titleStyle);
        excel.setCellValue(sheet, sheetrow, 1, "表名", titleStyle);
        excel.setCellValue(sheet, sheetrow, 2, "说明", titleStyle);
        int index =1;
        for (Table table : tables) {
            sheetrow++;
            excel.setCellLinkValue(sheet, sheetrow, 0, String.valueOf(index), table.getName() + "!A1");
            excel.setCellStyle(sheet, sheetrow, 0, defaultStyle);

            excel.setCellValue(sheet, sheetrow, 1, table.getName(), defaultStyle);
            excel.setCellValue(sheet, sheetrow, 2, table.getComment(), defaultStyle);

            generatorTable(table);
            index++;
        }
        excel.save(db + ".xls");

    }

    private void setColumnWidth(HSSFSheet sheet, int colIndex, int width) {
        sheet.setColumnWidth(colIndex, width * 256);
    }

    

    private void generatorTable(Table table) throws Exception {
        CellRangeAddress region = null;
        List<Column> columns = table.getColumns();
        //生成表头
        HSSFSheet sheet = excel.getSheet(table.getName(), true);
        sheet.setDefaultRowHeightInPoints(20);
        setColumnWidth(sheet, 0, 30);
        setColumnWidth(sheet, 1, 20);
        setColumnWidth(sheet, 2, 10);
        setColumnWidth(sheet, 3, 10);
        setColumnWidth(sheet, 4, 10);
        setColumnWidth(sheet, 5, 15);
        setColumnWidth(sheet, 6, 60);
        int sheetrow = 0;
        excel.setCellValue(sheet, sheetrow, 0, "表名", titleStyle);

        region = excel.addMergedRegion(sheet, sheetrow, sheetrow, 1, 6);
        excel.setRegionBorder(sheet, region, "1111", CellStyle.BORDER_THIN, HSSFColor.PALE_BLUE.index);
        excel.setCellValue(sheet, sheetrow, 1, table.getName(), defaultStyle);

        sheetrow++;
        excel.setCellValue(sheet, sheetrow, 0, "说明", titleStyle);

        region = excel.addMergedRegion(sheet, sheetrow, sheetrow, 1, 6);
        excel.setRegionBorder(sheet, region, "1111", CellStyle.BORDER_THIN, HSSFColor.PALE_BLUE.index);
        excel.setCellValue(sheet, sheetrow, 1, table.getComment(), defaultStyle);

        sheetrow++;
        excel.setCellValue(sheet, sheetrow, 0, "名称", titleStyle);
        excel.setCellValue(sheet, sheetrow, 1, "数据类型", titleStyle);
        excel.setCellValue(sheet, sheetrow, 2, "为空", titleStyle);
        excel.setCellValue(sheet, sheetrow, 3, "主键", titleStyle);
        excel.setCellValue(sheet, sheetrow, 4, "默认值", titleStyle);
        excel.setCellValue(sheet, sheetrow, 5, "自增", titleStyle);
        excel.setCellValue(sheet, sheetrow, 6, "备注", titleStyle);

        for (Column col : columns) {
            sheetrow++;
            excel.setCellValue(sheet, sheetrow, 0, col.getName(), defaultStyle);
            excel.setCellValue(sheet, sheetrow, 1, col.getType(), defaultStyle);
            excel.setCellValue(sheet, sheetrow, 2, col.getEmpty(), defaultStyle);
            excel.setCellValue(sheet, sheetrow, 3, col.getKey(), defaultStyle);
            excel.setCellValue(sheet, sheetrow, 4, col.getDefaultValue(), defaultStyle);
            excel.setCellValue(sheet, sheetrow, 5, col.getExtra(), defaultStyle);
            excel.setCellValue(sheet, sheetrow, 6, col.getComment(), defaultStyle);
        }
    }

}

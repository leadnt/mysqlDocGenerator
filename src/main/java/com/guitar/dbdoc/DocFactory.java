package com.guitar.dbdoc;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.Statement;
import java.util.ArrayList;
import java.util.List;
import java.util.Properties;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Hyperlink;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.RegionUtil;

/**
 *
 * @author hxy
 */
public class DocFactory {
    private Properties prop;
    Connection con = null;
    HSSFWorkbook book = null;
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
        try{
            book = new HSSFWorkbook();
            CreationHelper createHelper = book.getCreationHelper();
            titleFont = book.createFont();
            titleFont.setFontName("宋体");
            titleFont.setFontHeightInPoints((short)11);
            titleFont.setColor(HSSFColor.WHITE.index);

            defaultFont = book.createFont();
            defaultFont.setFontName("宋体");
            defaultFont.setFontHeightInPoints((short)11);

            titleStyle = book.createCellStyle();
            titleStyle.setFillForegroundColor(HSSFColor.ROYAL_BLUE.index);
            titleStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);
            titleStyle.setFont(titleFont);
            
            titleStyle.setBorderBottom(HSSFCellStyle.BORDER_THIN); //下边框
            titleStyle.setBottomBorderColor(HSSFColor.PALE_BLUE.index);
            titleStyle.setBorderLeft(HSSFCellStyle.BORDER_THIN);//左边框
            titleStyle.setLeftBorderColor(HSSFColor.PALE_BLUE.index);
            titleStyle.setBorderTop(HSSFCellStyle.BORDER_THIN);//上边框
            titleStyle.setTopBorderColor(HSSFColor.PALE_BLUE.index);
            titleStyle.setBorderRight(HSSFCellStyle.BORDER_THIN);//右边框
            titleStyle.setRightBorderColor(HSSFColor.PALE_BLUE.index);

            defaultStyle = book.createCellStyle();
            defaultStyle.setFont(defaultFont);
            defaultStyle.setBorderBottom(HSSFCellStyle.BORDER_THIN); //下边框
            defaultStyle.setBottomBorderColor(HSSFColor.PALE_BLUE.index);
            defaultStyle.setBorderLeft(HSSFCellStyle.BORDER_THIN);//左边框
            defaultStyle.setLeftBorderColor(HSSFColor.PALE_BLUE.index);
            defaultStyle.setBorderTop(HSSFCellStyle.BORDER_THIN);//上边框
            defaultStyle.setTopBorderColor(HSSFColor.PALE_BLUE.index);
            defaultStyle.setBorderRight(HSSFCellStyle.BORDER_THIN);//右边框
            defaultStyle.setRightBorderColor(HSSFColor.PALE_BLUE.index);
            
            
            Class.forName("com.mysql.jdbc.Driver");
            con = DriverManager.getConnection(url + db, user, pwd);
            HSSFSheet sheet = book.createSheet("目录");
            
            sheet.setDefaultRowHeightInPoints(20);
            
            setColumnWidth(sheet,0,10);
            setColumnWidth(sheet,1,40);
            setColumnWidth(sheet,2,60);
            List<Table> tables = queryTables();
            int sheetrow = 0;
            HSSFRow row = null;
            HSSFCell cell = null;
            row = getRow(sheet, sheetrow);
            cell = row.createCell(0,HSSFCell.CELL_TYPE_STRING);
            cell.setCellValue("操作");
            cell.setCellStyle(titleStyle);
            
            cell = row.createCell(1,HSSFCell.CELL_TYPE_STRING);
            cell.setCellValue("表名");
            cell.setCellStyle(titleStyle);
            
            cell = row.createCell(2,HSSFCell.CELL_TYPE_STRING);
            cell.setCellValue("说明");
            cell.setCellStyle(titleStyle);
            
            for(Table table:tables){
                sheetrow++;
                row = getRow(sheet, sheetrow);
                cell = row.createCell(0,HSSFCell.CELL_TYPE_STRING);
                Hyperlink link = createHelper.createHyperlink(Hyperlink.LINK_DOCUMENT);
                link.setAddress(table.getName() + "!A1");
                link.setLabel("打开");
                cell.setCellValue("打开");
                cell.setHyperlink(link);
                cell.setCellStyle(defaultStyle);
                
                cell = row.createCell(1,HSSFCell.CELL_TYPE_STRING);
                cell.setCellValue(table.getName());
                cell.setCellStyle(defaultStyle);
                
                cell = row.createCell(2,HSSFCell.CELL_TYPE_STRING);
                cell.setCellValue(table.getComment());
                cell.setCellStyle(defaultStyle);
                
                generatorTable(table);
            }
            FileOutputStream fopts = new FileOutputStream(db + ".xls");
            book.write(fopts);
            fopts.flush();
            fopts.close();
        }finally{
            if(con != null){
                con.close();
            }
        }
    }

    private HSSFRow getRow(HSSFSheet sheet, int sheetrow) {
        HSSFRow row = sheet.createRow(sheetrow);
        //row.setHeightInPoints(25);
        return row;
    }
    private void setColumnWidth(HSSFSheet sheet,int colIndex,int width){
        sheet.setColumnWidth(colIndex, width * 256);
    }
    
    private List<Table> queryTables() throws Exception {
        ArrayList<Table> tables = new ArrayList<Table>();
        Statement st = null;
        ResultSet rs = null;
        try {
            st = con.createStatement();
            rs = st.executeQuery("show table status where comment != 'VIEW';");
            while (rs.next()) {
                String tableName = rs.getString("Name");
                String comment = rs.getString("Comment");
                Table table = new Table();
                table.setName(tableName);
                table.setComment(comment);
                tables.add(table);
            }
        }finally{
            if(rs != null){
                rs.close();
            }
            if(st != null){
                st.close();
            }
        }
        return tables;
    }
    
    private List<Column> queryColumns(String tableName) throws Exception{
        ArrayList<Column> columns = new ArrayList<Column>();
        Statement st = null;
        ResultSet rs = null;
        try {
            st = con.createStatement();
            rs = st.executeQuery("show full columns from " + tableName);
            while (rs.next()) {
                Column column = new Column();
                column.setType(rs.getString("Type"));
                column.setName(rs.getString("Field"));
                column.setComment(rs.getString("Comment"));
                column.setEmpty(rs.getString("Null"));
                column.setKey(rs.getString("Key"));
                column.setDefaultValue(rs.getString("Default"));
                column.setExtra(rs.getString("Extra"));
                column.setComment(rs.getString("Comment"));
                
                columns.add(column);
            }
        }finally{
            if(rs != null){
                rs.close();
            }
            if(st != null){
                st.close();
            }
        }
        return columns;
    }

    private void generatorTable(Table table) throws Exception {
        CellRangeAddress region = null;
        List<Column> columns = queryColumns(table.getName());
        //生成表头
        HSSFSheet sheet = book.createSheet(table.getName());
        sheet.setDefaultRowHeightInPoints(20);
        setColumnWidth(sheet,0,30);
        setColumnWidth(sheet,1,20);
        setColumnWidth(sheet,2,10);
        setColumnWidth(sheet,3,10);
        setColumnWidth(sheet,4,10);
        setColumnWidth(sheet,5,15);
        setColumnWidth(sheet,6,60);
        int sheetrow = 0;
        HSSFRow row = null;
        HSSFCell cell = null;
        row = getRow(sheet, sheetrow);
        cell = row.createCell(0,HSSFCell.CELL_TYPE_STRING);
        cell.setCellValue("表名");
        cell.setCellStyle(titleStyle);
        
        region = new CellRangeAddress(sheetrow, sheetrow, (short) 1, (short) 6);
        setRegionBorder(region,sheet,book);
        sheet.addMergedRegion(region);
        cell = row.createCell(1,HSSFCell.CELL_TYPE_STRING);
        cell.setCellValue(table.getName());
        cell.setCellStyle(defaultStyle);
        sheetrow++;
        row = getRow(sheet, sheetrow);
        cell = row.createCell(0,HSSFCell.CELL_TYPE_STRING);
        cell.setCellValue("说明");
        cell.setCellStyle(titleStyle);
        region = new CellRangeAddress(sheetrow, sheetrow, (short) 1, (short) 6);
        sheet.addMergedRegion(region);
        setRegionBorder(region,sheet,book);
        cell = row.createCell(1,HSSFCell.CELL_TYPE_STRING);
        cell.setCellValue(table.getComment());
        cell.setCellStyle(defaultStyle);
        sheetrow++;
        row = getRow(sheet, sheetrow);
        cell = row.createCell(0,HSSFCell.CELL_TYPE_STRING);
        cell.setCellValue("名称");
        cell.setCellStyle(titleStyle);
        cell = row.createCell(1,HSSFCell.CELL_TYPE_STRING);
        cell.setCellValue("数据类型");
        cell.setCellStyle(titleStyle);
        cell = row.createCell(2,HSSFCell.CELL_TYPE_STRING);
        cell.setCellValue("为空");
        cell.setCellStyle(titleStyle);
        cell = row.createCell(3,HSSFCell.CELL_TYPE_STRING);
        cell.setCellValue("主键");
        cell.setCellStyle(titleStyle);
        cell = row.createCell(4,HSSFCell.CELL_TYPE_STRING);
        cell.setCellValue("默认值");
        cell.setCellStyle(titleStyle);
        cell = row.createCell(5,HSSFCell.CELL_TYPE_STRING);
        cell.setCellValue("自增");
        cell.setCellStyle(titleStyle);
        cell = row.createCell(6,HSSFCell.CELL_TYPE_STRING);
        cell.setCellValue("备注");
        cell.setCellStyle(titleStyle);
        for(Column col:columns){
            sheetrow++;
            row = getRow(sheet, sheetrow);
            cell = row.createCell(0,HSSFCell.CELL_TYPE_STRING);
            cell.setCellValue(col.getName());
            cell.setCellStyle(defaultStyle);
            cell = row.createCell(1,HSSFCell.CELL_TYPE_STRING);
            cell.setCellValue(col.getType());
            cell.setCellStyle(defaultStyle);
            cell = row.createCell(2,HSSFCell.CELL_TYPE_STRING);
            cell.setCellValue(col.getEmpty());
            cell.setCellStyle(defaultStyle);
            cell = row.createCell(3,HSSFCell.CELL_TYPE_STRING);
            cell.setCellValue(col.getKey());
            cell.setCellStyle(defaultStyle);
            cell = row.createCell(4,HSSFCell.CELL_TYPE_STRING);
            cell.setCellValue(col.getDefaultValue());
            cell.setCellStyle(defaultStyle);
            cell = row.createCell(5,HSSFCell.CELL_TYPE_STRING);
            cell.setCellValue(col.getExtra());
            cell.setCellStyle(defaultStyle);
            cell = row.createCell(6,HSSFCell.CELL_TYPE_STRING);
            cell.setCellValue(col.getComment());
            cell.setCellStyle(defaultStyle);
        }
    }
    private static void setRegionBorder(CellRangeAddress region, Sheet sheet,Workbook wb){  
        RegionUtil.setBorderBottom(HSSFCellStyle.BORDER_THIN,region, sheet, wb);  
        RegionUtil.setBorderLeft(HSSFCellStyle.BORDER_THIN,region, sheet, wb);  
        RegionUtil.setBorderRight(HSSFCellStyle.BORDER_THIN,region, sheet, wb);  
        RegionUtil.setBorderTop(HSSFCellStyle.BORDER_THIN,region, sheet, wb);  
        RegionUtil.setBottomBorderColor(HSSFColor.PALE_BLUE.index, region, sheet, wb);
        RegionUtil.setTopBorderColor(HSSFColor.PALE_BLUE.index, region, sheet, wb);
        RegionUtil.setLeftBorderColor(HSSFColor.PALE_BLUE.index, region, sheet, wb);
        RegionUtil.setRightBorderColor(HSSFColor.PALE_BLUE.index, region, sheet, wb);
    } 
    
}

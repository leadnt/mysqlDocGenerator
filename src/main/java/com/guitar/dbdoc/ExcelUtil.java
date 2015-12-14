package com.guitar.dbdoc;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Hyperlink;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.RegionUtil;

/**
 *
 * @author hxy
 */
public class ExcelUtil {
    private HSSFWorkbook book = null;
    private POIFSFileSystem fs;
    private File file;


    public ExcelUtil() {
        book = new HSSFWorkbook();
    }
    public ExcelUtil(File file) throws IOException{
        this.file = file;
        fs = new POIFSFileSystem(new FileInputStream(file));
        book = new HSSFWorkbook(fs);
    }
    public HSSFSheet getSheet(String name ,boolean create){
        HSSFSheet sheet = book.getSheet(name);
        if(sheet == null && create){
            sheet = book.createSheet(name);
        }
        return sheet;
    }
    public void save() throws Exception{
        if(file == null){
            return ;
        }
        try (FileOutputStream fopts = new FileOutputStream(file)) {
            book.write(fopts);
            fopts.flush();
        }
    }
    public void save(File file) throws Exception{
        try (FileOutputStream fopts = new FileOutputStream(file)) {
            book.write(fopts);
            fopts.flush();
        }
    }
    public void save(String filename) throws Exception{
        File file = new File(filename);
        try (FileOutputStream fopts = new FileOutputStream(file)) {
            book.write(fopts);
            fopts.flush();
        }
    }
    public HSSFRow getRow(String name,int rowIndex){
        HSSFSheet sheet = getSheet(name,true);
        return getRow(sheet,rowIndex);
    }
    public HSSFRow getRow(HSSFSheet sheet,int rowIndex){
        HSSFRow row = sheet.getRow(rowIndex);
        if(row == null){
            row = sheet.createRow(rowIndex);
        }
        return row;
    }
    public HSSFCell getCell(String name,int rowIndex,int colIndex){
        HSSFRow row = getRow(name,rowIndex);
        HSSFCell cell = row.getCell(colIndex);
        if(cell == null){
            cell = row.createCell(colIndex);
        }
        return cell;
    }
    public HSSFCell getCell(HSSFSheet sheet,int rowIndex,int colIndex){
        HSSFRow row = getRow(sheet,rowIndex);
        HSSFCell cell = row.getCell(colIndex);
        if(cell == null){
            cell = row.createCell(colIndex);
        }
        return cell;
    }
    
    public HSSFCell setCellValue(HSSFSheet sheet,int rowIndex,int colIndex,String value,int type){
        HSSFCell cell = getCell(sheet,rowIndex,colIndex);
        cell.setCellValue(value);
        return cell;
    }
    
    public HSSFCell setCellValue(HSSFSheet sheet,int rowIndex,int colIndex,String value){
        return setCellValue(sheet,rowIndex,colIndex,value,HSSFCell.CELL_TYPE_STRING);
    }
    public HSSFCell setCellValue(HSSFSheet sheet,int rowIndex,int colIndex,String value,CellStyle style){
        setCellValue(sheet,rowIndex,colIndex,value,HSSFCell.CELL_TYPE_STRING);
        return setCellStyle(sheet,rowIndex,colIndex,style);
    }
    public HSSFCell setCellLinkValue(HSSFSheet sheet,int rowIndex,int colIndex,String value,String link){
        HSSFCell cell = getCell(sheet,rowIndex,colIndex);
        Hyperlink hyperlink = getCreationHelper().createHyperlink(Hyperlink.LINK_DOCUMENT);
        cell.setCellValue(value);
        hyperlink.setLabel(value);
        hyperlink.setAddress(link);
        cell.setHyperlink(hyperlink);
        return cell;
    }
    
    public HSSFCell setCellStyle(HSSFSheet sheet,int rowIndex,int colIndex,CellStyle style){
        HSSFCell cell = getCell(sheet,rowIndex,colIndex);
        cell.setCellStyle(style);
        return cell;
    }
    public CreationHelper getCreationHelper(){
        return book.getCreationHelper();
    }
    public Font createFont(){
        return book.createFont();
    }
    public CellStyle createStyle(){
        return book.createCellStyle();
    }
    public CellRangeAddress addMergedRegion(HSSFSheet sheet,int firstRow, int lastRow, int firstCol, int lastCol){
        CellRangeAddress region = region = new CellRangeAddress(firstRow, lastRow, firstCol, lastCol);
        sheet.addMergedRegion(region);
        return region;
    }
    public CellStyle setStyleBorder(CellStyle style,String position,short border){
        if(position.charAt(0) == '1'){
            style.setBorderTop(border);
        }
        if(position.charAt(1) == '1'){
            style.setBorderBottom(border);
        }
        if(position.charAt(2) == '1'){
            style.setBorderLeft(border);
        }
        if(position.charAt(3) == '1'){
            style.setBorderRight(border);
        }
        return style;
    }
    public CellStyle setStyleBorderColor(CellStyle style,String position,short color){
        if(position.charAt(0) == '1'){
            style.setTopBorderColor(color);
        }
        if(position.charAt(1) == '1'){
            style.setBottomBorderColor(color);
        }
        if(position.charAt(2) == '1'){
            style.setLeftBorderColor(color);
        }
        if(position.charAt(3) == '1'){
            style.setRightBorderColor(color);
        }
        return style;
    }
    public void setRegionBorder(HSSFSheet sheet,CellRangeAddress region,String position,int border){
        if(position.charAt(0) == '1'){
            RegionUtil.setBorderTop(border,region, sheet, book); 
        }
        if(position.charAt(1) == '1'){
            RegionUtil.setBorderBottom(border,region, sheet, book);
        }
        if(position.charAt(2) == '1'){
            RegionUtil.setBorderLeft(border,region, sheet, book); 
        }
        if(position.charAt(3) == '1'){
            RegionUtil.setBorderRight(border,region, sheet, book);
        }
    }
    public void setRegionBorder(HSSFSheet sheet,CellRangeAddress region,String position,int border,int color){
        if(position.charAt(0) == '1'){
            RegionUtil.setBorderTop(border,region, sheet, book); 
            RegionUtil.setTopBorderColor(color,region, sheet, book); 
        }
        if(position.charAt(1) == '1'){
            RegionUtil.setBorderBottom(border,region, sheet, book);
            RegionUtil.setBottomBorderColor(color,region, sheet, book);
        }
        if(position.charAt(2) == '1'){
            RegionUtil.setBorderLeft(border,region, sheet, book); 
            RegionUtil.setLeftBorderColor(color,region, sheet, book);
        }
        if(position.charAt(3) == '1'){
            RegionUtil.setBorderRight(border,region, sheet, book);
            RegionUtil.setRightBorderColor(color,region, sheet, book);
        }
    }
    public void setRegionBorderColor(HSSFSheet sheet,CellRangeAddress region,String position,int color){
        if(position.charAt(0) == '1'){
            RegionUtil.setTopBorderColor(color,region, sheet, book); 
        }
        if(position.charAt(1) == '1'){
            RegionUtil.setBottomBorderColor(color,region, sheet, book);
        }
        if(position.charAt(2) == '1'){
            RegionUtil.setLeftBorderColor(color,region, sheet, book); 
        }
        if(position.charAt(3) == '1'){
            RegionUtil.setRightBorderColor(color,region, sheet, book);
        }
    }
    
    
}

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.util.CellReference;

import java.util.*;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.FileOutputStream;

public class test {
    public static void main(String[] args) throws IOException {
        BookExcel book=new BookExcel();
        book.newBook("testExcel.xls",book.checkBook("2.xls"));
    }

    public static  String getCell(Cell cell){
        String result="";
        switch (cell.getCellType()) {
            case STRING:
                result = cell.getRichStringCellValue().getString();
                break;
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    result = cell.getDateCellValue().toString();
                } else {
                    result = Integer.toString((int)cell.getNumericCellValue());
                }
                break;
        }
        return result;
    }
    public  static void style(Workbook wb, Cell cell, IndexedColors color, BorderStyle bord){
        CellStyle cellStyle = wb.createCellStyle();
        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        cellStyle.setBorderBottom(bord);
        cellStyle.setBorderLeft(bord);
        cellStyle.setBorderRight(bord);
        cellStyle.setBorderTop(bord);
        cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        cellStyle.setFillForegroundColor(color.getIndex());
        cell.setCellStyle(cellStyle);
    }
    public static void styleTable(Workbook wb,Cell cell,Row row){
        if(row.getRowNum()==0) {
            style(wb, cell,IndexedColors.BLUE_GREY,BorderStyle.MEDIUM);
        }else
        {
            style(wb,cell,IndexedColors.SKY_BLUE,BorderStyle.THIN);
        }
    }
    public  static void printCell(Cell cell){
        System.out.print(getCell(cell)+" ");
        //System.out.println();
    }
}

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
        book.newBook("/home/abdulovkt/Документы/ProjectFiles/testExcel.xls",book.checkBook("2.xls"));
        //Row row = sheet.createRow(0);
        //Cell cell=row.createCell(0);
        //Создание Map и ArrayList для заполнения книги;



        /*Workbook wb = WorkbookFactory.create(new FileInputStream(new File("2.xls")));
        Map<Integer, List<String>> rowSheet= new HashMap<Integer,List<String>>();
        for (Row row:wb.getSheetAt(0)) {
            ArrayList<String> cellVal = new ArrayList();
            for (int i=0; i<row.getLastCellNum(); i++) {
                if(i==2) {
                    if (row.getCell(i).getNumericCellValue() == 1.0) {
                        cellVal.add(row.getCell(i).getStringCellValue());
                    }
                }
            }
            rowSheet.put(row.getRowNum(),cellVal);
        }*/
            /*for (Cell cell:row) {
                printCell(cell);
                styleTable(wb,cell,row);
                if(row.getCell(2).getCellType()==CellType.NUMERIC) {
                    if (row.getCell(2).getNumericCellValue() == 1.0) {
                        style(wb,  cell, IndexedColors.RED1, BorderStyle.THIN);
                    }
                }
            }
            System.out.println();
        }
        FileOutputStream file = new FileOutputStream("2.xls");
        wb.write(file);
        wb.close();
        file.close();*/
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

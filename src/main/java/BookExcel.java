import org.apache.poi.ss.usermodel.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.text.DateFormat;
import java.util.*;

public class BookExcel {
    String name;
    String nameSheet;
    Map<Integer, List<String>> filterSheet;

    public BookExcel() {
        filterSheet= new LinkedHashMap<Integer, List<String>>();
    }

    public void setName(String name) {
        this.name = name;
    }

    public void newBook(String name, List list) throws IOException {
        int countRow = 1;
        Map<Integer, List<String>> mapCheckBook = new HashMap<Integer, List<String>>();
        List<String> valCell;
        Calendar thisDay = new GregorianCalendar();
        HSSFWorkbook wb = new HSSFWorkbook();
        nameSheet = "Трудозатраты за " + getWeekDate(thisDay);
        System.out.println(nameSheet);
        Sheet sheet = wb.createSheet(nameSheet);
        //тест по записи Map в книгу
        for(int i=0;i<list.size();i++) {
            Map<Integer,List<String>>writeCells=(Map)list.get(i);
            for (Map.Entry<Integer, List<String>> valueCell : writeCells.entrySet()) {
                Row row = sheet.createRow(countRow);
                valCell = new ArrayList<String>(valueCell.getValue());
                for (int j = 0; j < valCell.size(); j++) {
                    Cell cell = row.createCell(j);
                    cell.setCellValue(valCell.get(j));
                }
                countRow++;
            }
        }
        //--------------------------
        FileOutputStream file = new FileOutputStream(name);
        wb.write(file);
        wb.close();
    }

    public String getWeekDate(Calendar date) {
        SimpleDateFormat dateFormat = new SimpleDateFormat("dd.MM");
        String begWeek = "";
        String endWeek = "";
        endWeek = dateFormat.format(date.getTime());
        date.add(Calendar.DATE, -7);
        begWeek = dateFormat.format(date.getTime());
        String weekDate = begWeek + "-" + endWeek;
        return weekDate;
    }

    public void style(Workbook wb, Cell cell, IndexedColors color, BorderStyle bord) {
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

    public void styleTable(Workbook wb, Cell cell, Row row) {
        if (row.getRowNum() == 0) {
            style(wb, cell, IndexedColors.BLUE_GREY, BorderStyle.MEDIUM);
        } else {
            style(wb, cell, IndexedColors.SKY_BLUE, BorderStyle.THIN);
        }
    }

    public List checkBook(String name) throws IOException {
        List<Map>dateSheet=new ArrayList<Map>();
        Workbook wb = WorkbookFactory.create(new FileInputStream(new File(name)));
        //Sheet sheet=wb.getSheetAt(0);
        bookFilter(wb);
        dateSheet.add(filterSheet);
        /*Map<Integer, List<String>> rowSheet = new HashMap<Integer, List<String>>();
        for (Row row : wb.getSheetAt(0)) {
            List<String> cellVal = new ArrayList();
            for (int i = 0; i < row.getLastCellNum(); i++) {
                if (row.getCell(2, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL)==null) {
                    cellVal.add(getCell(row.getCell(i)));
                }
            }
            if(cellVal.size() >0) {
                rowSheet.put(row.getRowNum(), cellVal);
            }
        }
        dateSheet.add(rowSheet);*/
        return dateSheet;
    }
    public void bookFilter(Workbook wb){
        //Map<Integer, List<String>> filterSheet = new HashMap<Integer, List<String>>();
        Sheet sheet=wb.getSheetAt(0);
        sheetFind(sheet,"Календарный день",12);
        sheetFind(sheet,"Сотрудники",0);
        sheetFind(sheet, "Результат", 2);

        /*for (Row row : wb.getSheetAt(0)) {
            List<String> cellVal = new ArrayList<String>();
            for (int i = 0; i < row.getLastCellNum(); i++) {
                if( isString(row.getCell(2))) {
                    if (row.getCell(2).getStringCellValue().equals("Результат")) {
                        cellVal.add(String.valueOf(getCell(row.getCell(i))));
                    }
                }
            }
            if(cellVal.size() >0) {
                filterSheet.put(row.getRowNum(), cellVal);
            }
        }*/
        //return filterSheet;
    }

    public void sheetFind(Sheet sheet, String name, int numCell){


        for(int j=0;j<sheet.getLastRowNum();j++) {
            List<String> cellVal = new ArrayList<String>();
            Row row = sheet.getRow(j);
            if(isNullCell(row.getCell(numCell))!=true) {
                DateFormat dateFormat = new SimpleDateFormat("dd.MM");
                for (int i = 0; i < row.getLastCellNum(); i++) {
                    if (isString(row.getCell(numCell))) {
                        if (row.getCell(numCell).getStringCellValue().equals(name)) {
                           if(row.getCell(i).getCellType() == CellType.NUMERIC) {
                               if(DateUtil.isCellDateFormatted(row.getCell(i))){
                                   cellVal.add(String.valueOf(dateFormat.format(row.getCell(i).getDateCellValue())));
                               }else{
                                cellVal.add(String.format("%.2f",row.getCell(i).getNumericCellValue()));
                               }
                            }else{
                               cellVal.add(row.getCell(i).getStringCellValue());
                           }
                        }
                    }
                }
            }
            if(cellVal.size()>0) {
                for(int l=0;l<10;l++){
                    cellVal.remove(3);
                }
                filterSheet.put(row.getRowNum(), cellVal);
            }
        }
    }
    public String getCell(Cell cell) {
        String result = "";
        switch (cell.getCellType()) {
            case STRING:
                result = cell.getRichStringCellValue().getString();
                break;
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    result = cell.getDateCellValue().toString();
                } else {
                    result = Double.toString(cell.getNumericCellValue());
                }
                break;
        }
        return result;
    }

    public boolean isString(Cell cell) {
        boolean check=false;
        try {
            if (cell.getCellType() == CellType.STRING) {
                check=true;
            }
        }catch (NullPointerException e){
            return false;
        }
        return check;
    }
    public  boolean isNullCell(Cell cell){
        if(cell == null || cell.getCellType() == CellType.BLANK){
            return true;
        }else
            return false;
    }

}


import org.apache.poi.ss.usermodel.*;
import org.apache.poi.hssf.usermodel.*;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.*;

public class BookExcel {
    String name;
    String nameSheet;

    public BookExcel() {
    }

    public void setName(String name) {
        this.name = name;
    }

    public void newBook(String name, Map map) throws IOException {
        int countRow = 0;
        Map<Integer, List<String>> mapCheckBook = new HashMap<Integer, List<String>>();
        mapCheckBook = map;
        List<String> valCell;
        Calendar thisDay = new GregorianCalendar();
        HSSFWorkbook wb = new HSSFWorkbook();
        nameSheet = "Трудозатраты за " + getWeekDate(thisDay);
        System.out.println(nameSheet);
        Sheet sheet = wb.createSheet(nameSheet);
        //style(wb,);
        //тест по записи Map в книгу
        for (Map.Entry<Integer, List<String>> valueCell : mapCheckBook.entrySet()) {
            Row row = sheet.createRow(countRow);
            valCell = new ArrayList<String>(valueCell.getValue());
            for (int i = 0; i < valCell.size(); i++) {
                Cell cell = row.createCell(i);
                cell.setCellValue(valCell.get(i));
            }
            countRow++;
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

    public Map checkBook(String name) throws IOException {
        Workbook wb = WorkbookFactory.create(new FileInputStream(new File(name)));
        Map<Integer, List<String>> rowSheet = new HashMap<Integer, List<String>>();
        for (Row row : wb.getSheetAt(0)) {
            List<String> cellVal = new ArrayList();
            for (int i = 0; i < row.getLastCellNum(); i++) {
                if (row.getCell(2, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL)==null) {
                    cellVal.add(getCell(row.getCell(i)));}
                }
                if(cellVal.size() >0) {
                rowSheet.put(row.getRowNum(), cellVal);
                }
            }

        return rowSheet;
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
                    result = Integer.toString((int) cell.getNumericCellValue());
                }
                break;
        }
        return result;
    }

    public boolean isInteger(Cell cell) {
        boolean check=false;
        if (cell.getCellType() == CellType.NUMERIC) {
            check=true;
        }
        return check;
    }
}


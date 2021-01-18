import org.apache.poi.ss.usermodel.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellUtil;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.text.DateFormat;
import java.util.*;

public class BookExcel {
    String name;
    String nameSheet;
    Map<Integer, List<Object>> filterSheet;

    public BookExcel() {
        filterSheet= new LinkedHashMap<Integer, List<Object>>();
    }

    public void setName(String name) {
        this.name = name;
    }

    public void newBook(String name, List list) throws IOException {
        int countRow = 1;
        //Map<Integer, List<Object>> mapCheckBook = new HashMap<Integer, List<Object>>();
        List<Object> valCell;
        Calendar thisDay = new GregorianCalendar();
        HSSFWorkbook wb = new HSSFWorkbook();
        nameSheet = "Трудозатраты за " + getWeekDate(thisDay);
        System.out.println(nameSheet);
        Sheet sheet = wb.createSheet(nameSheet);
        //тест по записи Map в книгу
        for(int i=0;i<list.size();i++) {
            Map<Integer,List<Object>>writeCells=(Map)list.get(i);
            for (Map.Entry<Integer, List<Object>> valueCell : writeCells.entrySet()) {
                Row row = sheet.createRow(countRow);
                valCell = new ArrayList<Object>(valueCell.getValue());
                for (int j = 0; j < valCell.size(); j++) {
                    Cell cell = row.createCell(j);
                    if(valCell.get(j) instanceof String){
                        cell.setCellValue((String)valCell.get(j));
                    }else if(valCell.get(j) instanceof Double){
                        cell.setCellValue((Double)valCell.get(j));
                    }
                }
                countRow++;
            }
            formattedBook(sheet);
            checkWrongTime(sheet);
            Row rowNotes=sheet.createRow(sheet.getLastRowNum()+2);
            Row rowNotes1=sheet.createRow(sheet.getLastRowNum()+1);
            Cell cell =rowNotes.createCell(0);
            Cell cell1 =rowNotes1.createCell(0);
            cellColor(sheet,cell,IndexedColors.RED.getIndex());
            cellColor(sheet,cell1,IndexedColors.YELLOW.getIndex());
            CellUtil.createCell(rowNotes,1,"Неверно списано время");
            CellUtil.createCell(rowNotes1,1,"отпуск");
        }
        //--------------------------
        FileOutputStream file = new FileOutputStream(name);
        wb.write(file);
        wb.close();
    }
    public void formattedBook(Sheet sheet){
        sheet.addMergedRegion(new CellRangeAddress(1,1,0,2));
        Row row = sheet.getRow(1);
        Cell cell = row.getCell(0);
        cell.setCellValue("Календарный день");
        sheet.addMergedRegion(new CellRangeAddress(2,2,0,1));
        CellStyle cellStyle=sheet.getWorkbook().createCellStyle();
        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        cellStyle.setBorderBottom(BorderStyle.THIN);
        cellStyle.setBorderLeft(BorderStyle.THIN);
        cellStyle.setBorderRight(BorderStyle.THIN);
        cellStyle.setBorderTop(BorderStyle.THIN);
        cellStyle.setWrapText(true);
        for(int j=0;j<3;j++){
            sheet.autoSizeColumn(j);
        }
        for (Row row1 : sheet.getWorkbook().getSheetAt(0)) {
            if(row1.getRowNum()==1||row1.getRowNum()==2){
                row1.setHeight((short) 700);
            }
            for (int i = 0; i < row.getLastCellNum(); i++) {
                cell = row1.getCell(i);
                cell.setCellStyle(cellStyle);
                if(i>2){
                    sheet.setColumnWidth(i,3350);
                }
            }
        }

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
        Sheet sheet=wb.getSheetAt(0);
        sheetFind(sheet,"Календарный день",12);
        sheetFind(sheet,"Сотрудники",0);
        sheetFind(sheet, "Результат", 2);
    }
public void cellColor(Sheet sheet, Cell cell, short indexColor){
    CellStyle cellColor=sheet.getWorkbook().createCellStyle();
    cellColor.setFillPattern(FillPatternType.SOLID_FOREGROUND);
    cellColor.setFillForegroundColor(indexColor);
    cell.setCellStyle(cellColor);
}
    public void checkWrongTime(Sheet sheet){
        Map<Integer,List<Double>>employes=new HashMap<Integer, List<Double>>();
        List<Double>timeChange= Arrays.asList(4.00,7.00,11.00,8.00);
        List<Double>restTime=Arrays.asList(8.25,8.0,7.00);
        List<Double>substitution=Arrays.asList(4.00,7.00,11.00,8.25);
        employes.put(3701140,timeChange);
        employes.put(3701277,timeChange);
        employes.put(3703077,timeChange);
        employes.put(3703422,timeChange);
        employes.put(3703149,substitution);
        employes.put(3703672,substitution);
        employes.put(3701146,restTime);
        employes.put(3701916,restTime);
        employes.put(3703413,restTime);
        DataFormatter formatter = new DataFormatter();
        for(int i=3;i<=sheet.getLastRowNum();i++){
            Row row=sheet.getRow(i);
            for(Map.Entry<Integer,List<Double>>mapEmploye: employes.entrySet()){
                if(Integer.parseInt(formatter.formatCellValue(row.getCell(0)))==mapEmploye.getKey()){
                    ArrayList<Double>emploeysTime=new ArrayList<Double>(mapEmploye.getValue());
                    for(int j=3;j<11;j++) {
                        if (!isString(row.getCell(j))){
                            int count=0;
                            for (double emploey:emploeysTime) {
                                    if(emploey!=(Double)getCell(row.getCell(j))){
                                        count++;
                                    }
                                    if(count==emploeysTime.size()){
                                        cellColor(sheet, row.getCell(j), IndexedColors.RED.getIndex());
                                    }
                            }
                            if(row.getCell(j).getNumericCellValue()==8.00){
                                cellColor(sheet, row.getCell(j), IndexedColors.YELLOW.getIndex());
                            }
                        }
                    }
                }


            }
        }

    }

    public void sheetFind(Sheet sheet, String name, int numCell){


        for(int j=0;j<sheet.getLastRowNum();j++) {
            List<Object> cellVal = new ArrayList<Object>();
            Row row = sheet.getRow(j);
            if(isNullCell(row.getCell(numCell))!=true) {
                for (int i = 0; i < row.getLastCellNum(); i++) {
                    if (isString(row.getCell(numCell))) {
                        if (row.getCell(numCell).getStringCellValue().equals(name)) {
                            cellVal.add(getCell(row.getCell(i)));
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
    public Object getCell(Cell cell) {
        Object result = "";
        DecimalFormat dFormat= new DecimalFormat("#,##");
        SimpleDateFormat dateFormat=new SimpleDateFormat("dd.MM");
        switch (cell.getCellType()) {
            case STRING:
                result = cell.getRichStringCellValue().getString();
                break;
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    result=String.valueOf(dateFormat.format(cell.getDateCellValue()));
                } else {
                    //result=String.format("%.2f",cell.getNumericCellValue());
                    result=Math.round(cell.getNumericCellValue()*100)/100.0;
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


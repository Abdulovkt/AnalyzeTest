import org.apache.poi.ss.usermodel.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellUtil;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.*;

public class BookExcel {
    //String name;
    String nameSheet;
    Map<Integer, List<Object>> filterSheet;
    Map<String,List<Object>> projects;

    public BookExcel() {
        filterSheet= new LinkedHashMap<Integer, List<Object>>();
        projects=new HashMap<String, List<Object>>();
    }

    public void newBook(String name, List list) throws IOException {
        int countRow = 1;
        List<Object> valCell;
        Calendar thisDay = new GregorianCalendar();
        XSSFWorkbook wb = new XSSFWorkbook();
        nameSheet = "Трудозатраты за " + getWeekDate(thisDay);
        System.out.println(nameSheet);
        Sheet sheet = wb.createSheet(nameSheet);
        FormatingExcelBook formatExcel = new FormatingExcelBook(sheet);
        //TODO сделать отдельный метод по записи в лист книги.
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
            formatExcel.formatListExcel();
            checkWrongTime(sheet);
            Row rowNotes=sheet.createRow(sheet.getLastRowNum()+2);
            Row rowNotes1=sheet.createRow(sheet.getLastRowNum()+1);
            Cell cell =rowNotes.createCell(0);
            Cell cell1 =rowNotes1.createCell(0);
            cellColor(sheet,cell,IndexedColors.RED.getIndex());
            cellColor(sheet,cell1,IndexedColors.YELLOW.getIndex());
            CellUtil.createCell(rowNotes,1,"Неверно списано время");
            CellUtil.createCell(rowNotes1,1,"отпуск");
            Row rowNameAnalize=sheet.createRow(sheet.getLastRowNum()+2);
            Cell cellNameAnalize=rowNameAnalize.createCell(0);
            sheet.addMergedRegion(new CellRangeAddress(rowNameAnalize.getRowNum(),rowNameAnalize.getRowNum(),0,1));
            cellNameAnalize.setCellValue("Анализ трудозатрат по проектам");
            countRow=sheet.getLastRowNum()+1;
            for (Map.Entry<String, List<Object>> project : projects.entrySet()) {
                Row row = sheet.createRow(countRow);
                valCell = new ArrayList<Object>(project.getValue());
                for (int j = 0; j < valCell.size(); j++) {
                    Cell cellAnalize = row.createCell(j);
                    if(valCell.get(j) instanceof String){
                        cellAnalize.setCellValue((String)valCell.get(j));
                    }else if(valCell.get(j) instanceof Double){
                        cellAnalize.setCellValue((Double)valCell.get(j));
                    }
                    sheet.autoSizeColumn(0);
                    cellAnalize.setCellStyle(formatExcel.customizeStyleSheet());
                }
                countRow++;
            }
        }
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

    public void effortAnalysis(Sheet sheet){
            for (Row row:sheet.getWorkbook().getSheetAt(0)) {
                List<Object>project=new ArrayList<Object>();
                if(row.getRowNum()>4&&row.getCell(7).getStringCellValue().equals("Результат")!=true) {
                    if (!isNullCell(row.getCell(4))&&row.getCell(4).getStringCellValue().equals("Результат")!=true
                            &&row.getCell(4).getStringCellValue().equals("#")!=true) {
                        switch (Integer.parseInt(row.getCell(4).getStringCellValue())) {
                            case 746:
                                if (row.getCell(7).getStringCellValue().equals("Не присвоено")) {
                                    project.add(getCell(row.getCell(3)));
                                    project.add(getCell(row.getCell(1)));
                                    project.add(getCell(row.getCell(4)));
                                    project.add(getCell(row.getCell(6)));
                                    project.add(getCell(row.getCell(7)));
                                }
                                break;
                            case 748:
                                if (row.getCell(7).getStringCellValue().equals("Не присвоено")) {
                                    project.add(getCell(row.getCell(3)));
                                    project.add(getCell(row.getCell(1)));
                                    project.add(getCell(row.getCell(4)));
                                    project.add(getCell(row.getCell(6)));
                                    project.add(getCell(row.getCell(7)));
                                }
                                break;
                            case 750:
                                if (row.getCell(7).getStringCellValue().equals("Не присвоено")) {
                                    project.add(getCell(row.getCell(3)));
                                    project.add(getCell(row.getCell(1)));
                                    project.add(getCell(row.getCell(4)));
                                    project.add(getCell(row.getCell(6)));
                                    project.add(getCell(row.getCell(7)));
                                }
                                break;
                            case 745:
                                if (row.getCell(6).getStringCellValue().equals("Газпром")) {
                                    project.add(getCell(row.getCell(3)));
                                    project.add(getCell(row.getCell(1)));
                                    project.add(getCell(row.getCell(4)));
                                    project.add(getCell(row.getCell(6)));
                                    project.add(getCell(row.getCell(7)));
                                }
                                break;
                        }
                        if (project.size() > 0) {
                            projects.put((String)getCell(row.getCell(3)), project);
                            //count++;
                        }
                    }
                }
            }
    }

    public List checkBook(String name) throws IOException {
        List<Map>dateSheet=new ArrayList<Map>();
        Workbook wb = WorkbookFactory.create(new FileInputStream(new File(name)));
        effortAnalysis(wb.getSheetAt(0));
        bookFilterTime(wb);
        dateSheet.add(filterSheet);
        return dateSheet;
    }

    public void bookFilterTime(Workbook wb){
        Sheet sheet=wb.getSheetAt(0);
        sheetFind(sheet,"Календарный день",12);
        sheetFind(sheet,"Сотрудники",0);
        sheetFind(sheet, "Результат", 2);
    }

    public void cellColor(Sheet sheet, Cell cell, short indexColor){
    CellStyle cellColor=sheet.getWorkbook().createCellStyle();
    cellColor.setFillPattern(FillPatternType.SOLID_FOREGROUND);
    cellColor.setFillForegroundColor(indexColor);
    cellColor.setAlignment(HorizontalAlignment.CENTER);
    cellColor.setVerticalAlignment(VerticalAlignment.CENTER);
    cellColor.setBorderBottom(BorderStyle.THIN);
    cellColor.setBorderLeft(BorderStyle.THIN);
    cellColor.setBorderRight(BorderStyle.THIN);
    cellColor.setBorderTop(BorderStyle.THIN);
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
                            for (double emploeyTime:emploeysTime) {
                                    if(emploeyTime!=(Double)getCell(row.getCell(j))){
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
            if(isNullRow(row)!=true&&(isNullCell(row.getCell(numCell))!=true)) {
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
        SimpleDateFormat dateFormat=new SimpleDateFormat("dd.MM");
        switch (cell.getCellType()) {
            case STRING:
                result = cell.getRichStringCellValue().getString();
                break;
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    result=String.valueOf(dateFormat.format(cell.getDateCellValue()));
                } else {
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
    public boolean isNullRow(Row row){
        if(row==null){
            return true;
        }else
            {
            return false;
        }
    }

}


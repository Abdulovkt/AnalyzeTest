import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;


public class FormatingExcelBook {
    Sheet sheet;
    public FormatingExcelBook(Sheet sheet){
       this.sheet=sheet;
    }
    public void formatListExcel(){
        sheet.addMergedRegion(new CellRangeAddress(1,1,0,2));
        Row row = sheet.getRow(1);
        Cell cell = row.getCell(0);
        cell.setCellValue("Календарный день");
        sheet.addMergedRegion(new CellRangeAddress(2,2,0,1));
        for(int j=0;j<3;j++){
            sheet.autoSizeColumn(j);
        }
        for (Row row1 : sheet.getWorkbook().getSheetAt(0)) {
            if(row1.getRowNum()==1||row1.getRowNum()==2){
                row1.setHeight((short) 700);
            }
            for (int i = 0; i < row.getLastCellNum(); i++) {
                cell = row1.getCell(i);
                cell.setCellStyle(customizeStyleSheet());
                if(i>2){
                    sheet.setColumnWidth(i,3350);
                }
            }
        }
    }
    CellStyle customizeStyleSheet(){
        CellStyle cellStyle=sheet.getWorkbook().createCellStyle();
        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        cellStyle.setBorderBottom(BorderStyle.THIN);
        cellStyle.setBorderLeft(BorderStyle.THIN);
        cellStyle.setBorderRight(BorderStyle.THIN);
        cellStyle.setBorderTop(BorderStyle.THIN);
        cellStyle.setWrapText(true);
        return cellStyle;
    }
}

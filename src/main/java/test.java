import java.io.IOException;


public class test {
    public static void main(String[] args) throws IOException {
        BookExcel book=new BookExcel();
        book.newBook("testExcel.xlsx",book.checkBook("0ANALYSIS_PATTERN.xls"));
    }
}

package readwritefromexcelfile;

import java.io.File;
import java.io.FileOutputStream;

import java.util.Map;
import java.util.Set;
import java.util.TreeMap;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WriteExcelFileFromUserInput {
    public static void main(String[] args) throws Exception {

        //Create blank workbook
        XSSFWorkbook workbook = new XSSFWorkbook();

        //Create a blank sheet
        XSSFSheet spreadsheet = workbook.createSheet( " Student Info ");

        //Create row object
        XSSFRow row;

        //This data needs to be written (Object[])
        Map < String, Object[] > studentinfo = new TreeMap < String, Object[] >();
        studentinfo.put( "1", new Object[] {
                "Name", "Roll", "ID" , "Dept" });

        studentinfo.put( "2", new Object[] {
                "Prince", "10", "112233", "CSE" });

        studentinfo.put( "3", new Object[] {
                "Imrul", "100","112244", "CSE" });

        //Iterate over data and write to sheet
        Set < String > keyid = studentinfo.keySet();
        int rowid = 0;

        for (String key : keyid) {
            row = spreadsheet.createRow(rowid++);
            Object [] objectArr = studentinfo.get(key);
            int cellid = 0;

            for (Object obj : objectArr){
                Cell cell = row.createCell(cellid++);
                cell.setCellValue((String)obj);
            }
        }
        //Write the workbook in file system
        FileOutputStream out = new FileOutputStream(
                new File("C:\\Users\\Progoti\\IdeaProjects\\JavaCodes\\src\\readwritefromexcelfile\\test2.xlsx"));

        workbook.write(out);
        out.close();
        System.out.println("test2.xlsx written successfully");
    }

}

import org.apache.poi.ss.usermodel.*;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

public class WriteExcel {
    public static void main(String[] args) {

        String path = "src/main/resources/users.xlsx";
        Workbook workbook = null;
        FileInputStream fileInputStream = null;

        try {
            fileInputStream = new FileInputStream(path);
            workbook = WorkbookFactory.create(fileInputStream);
        } catch (Exception e) {
            e.printStackTrace();
        }

        assert workbook != null;
        Sheet sheet = workbook.createSheet("students");

        List<String> students = new ArrayList<>(Arrays.asList("42","30","22","45","65","12"));

        for (int i = 0; i < students.size(); i++) {
            Row row = sheet.createRow(i);
            Cell cell = row.createCell(0);
            cell.setCellValue(Integer.parseInt(students.get(i)));
        }

        FileOutputStream fileOutputStream;

        try {
            fileInputStream.close();
            fileOutputStream = new FileOutputStream(path);
            workbook.write(fileOutputStream);
            workbook.close();
            fileOutputStream.close();
        } catch (IOException ex) {
            ex.printStackTrace();
        }


    }
}

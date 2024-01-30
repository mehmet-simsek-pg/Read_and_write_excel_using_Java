import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

public class ReadExcel {
    public static void main(String[] args) {

        String path = "src/main/resources/users.xlsx";
        String sheetName = "user";

        int columnCount = 2;
        int startRowNumber = 1;

        List<List<String>> users = new ArrayList<>();

        Workbook workbook = null;

        try {
            FileInputStream fileInputStream = new FileInputStream(path);
            workbook = WorkbookFactory.create(fileInputStream);
        } catch (IOException ex) {
            ex.printStackTrace();
        }

        assert workbook != null;
        Sheet sheet = workbook.getSheet(sheetName);
        int rowCount = sheet.getPhysicalNumberOfRows();

        for (int i = startRowNumber; i < rowCount ; i++) {
            List<String> rowList = new ArrayList<>();
            Row row = sheet.getRow(i);

            int cellCount = row.getPhysicalNumberOfCells();
            if (columnCount > cellCount){
                columnCount = cellCount;
            }

            for (int j = 0; j < columnCount; j++) {
                rowList.add(row.getCell(j).toString());
            }

            users.add(rowList);

        }

        System.out.println(Arrays.toString(users.toArray()));
    }
}

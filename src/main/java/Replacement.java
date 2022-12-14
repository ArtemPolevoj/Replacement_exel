import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.Iterator;

public class Replacement {
    static String replace( File file, String worldReplace, String replacement) {

        String txt;
        String outText = "";
        String nameFile = file.getName();
        int amount = 0;

            try (XSSFWorkbook workbook = new XSSFWorkbook(new FileInputStream(file))) {

                for (int s = 0; s < workbook.getNumberOfSheets(); s++) {
                    Sheet sheet = workbook.getSheetAt(s);

                    for (Row row : sheet) {

                        Iterator<Cell> cellIterator = row.cellIterator();

                        while (cellIterator.hasNext()) {

                            Cell cell = cellIterator.next();

                            if (cell.toString().contains(worldReplace)) {
                                cell.setCellValue(cell.toString().replaceAll(worldReplace, replacement));
                                amount++;
                            }
                        }
                    }
                }

                try (FileOutputStream writeFile = new FileOutputStream(file)) {
                    workbook.write(writeFile);
                }
                workbook.close();

                switch (amount) {
                    case (0):
                        txt = "? ????? \"" + nameFile + "\" ?????????? ?? ???????.\n";
                        break;
                    case (1):
                    case (21):
                    case (31):
                        txt = "? ????? \"" + nameFile + "\" ??????????? " + amount + " ??????.\n";
                        break;
                    case (2):
                    case (3):
                    case (4):
                    case (22):
                    case (23):
                    case (24):
                        txt = "? ????? \"" + nameFile + "\" ??????????? " + amount + " ??????.\n";
                        break;
                    default:
                        txt = "? ????? \"" + nameFile + "\" ??????????? " + amount + " ?????.\n";
                        break;
                }

                outText += txt;

            } catch (Exception e) {
                outText = "? ????? \"" + nameFile + "\" ?? ??????? ????????? ??????. ???? ?????. ???????? ?????? ? ?????? ?????????.\n";
            }
        return outText;
    }
}

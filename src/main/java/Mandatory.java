import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.*;
import java.io.*;
import java.util.Iterator;

public class Mandatory {
    static void mandatory() {

        File[] openFile = Files.open();
        if (openFile[0] == null) return;

        String fileName = Files.save();

        if (fileName.equals("")) return;

        boolean fileExists = false;

        if (new File(fileName).exists()) {

            try(XSSFWorkbook saveBook = new XSSFWorkbook(new FileInputStream(fileName)) ) {

                write(fileName, true,saveBook, openFile);

            } catch (IOException e) {
                JOptionPane.showMessageDialog(null, "Не удалось прочитать файл.");

            }

        } else {

            try (XSSFWorkbook saveBook = new XSSFWorkbook()){
                write(fileName + ".xlsx", false,saveBook, openFile);
            } catch (IOException e) {
                JOptionPane.showMessageDialog(null, "Не удалось создать новый файл.");

            }

        }
    }

    static void write(String fileName, boolean fileExists, XSSFWorkbook saveBook, File[] openFile)  {

        boolean exception = false;

        for (File file : openFile) {

            try (XSSFWorkbook workbook = new XSSFWorkbook(new FileInputStream(file))) {

                XSSFSheet openSheet = workbook.getSheet("Ремонт");

                if (openSheet == null || !file.getName().contains("ver")) {
                    JOptionPane.showMessageDialog(null, "В файле " + file.getName() + " нет ОЗ.\nВыберете файл(ы) с ver.2 и выше");
                    return;
                } else {

                    String sheetName = JOptionPane.showInputDialog(null, "Ведите имя листа", file.getName().substring(2));

                    XSSFSheet newSheet = saveBook.createSheet(sheetName.trim());
                    int rows = 7;
                    int colPart = 0;
                    int colQuantry = 1;
                    int colName = 2;
                    int colLabor = 0;

                    CellStyle borders = saveBook.createCellStyle();
                    borders.setBorderLeft(BorderStyle.THIN);
                    borders.setBorderRight(BorderStyle.THIN);
                    borders.setBorderBottom(BorderStyle.THIN);
                    borders.setBorderTop(BorderStyle.THIN);

                    Font font = saveBook.createFont();
                    font.setFontHeightInPoints((short) 12);
                    font.setBold(true);
                    CellStyle styleFont = saveBook.createCellStyle();
                    styleFont.setFont(font);
                    styleFont.setBorderLeft(BorderStyle.THIN);
                    styleFont.setBorderRight(BorderStyle.THIN);
                    styleFont.setBorderBottom(BorderStyle.THIN);
                    styleFont.setBorderTop(BorderStyle.THIN);

                    CellStyle center = saveBook.createCellStyle();
                    center.setBorderLeft(BorderStyle.THIN);
                    center.setBorderRight(BorderStyle.THIN);
                    center.setBorderBottom(BorderStyle.THIN);
                    center.setBorderTop(BorderStyle.THIN);
                    center.setAlignment(HorizontalAlignment.CENTER);

                    XSSFRow row1 = newSheet.createRow(0);
                    XSSFRow row2 = newSheet.createRow(2);
                    XSSFRow row3 = newSheet.createRow(3);
                    XSSFRow row4 = newSheet.createRow(4);
                    XSSFRow row5 = newSheet.createRow(6);
                    newSheet.setColumnWidth(0, 4000);
                    newSheet.setColumnWidth(1, 2000);
                    newSheet.setColumnWidth(2, 7000);

                    XSSFCell partTemp = newSheet.getRow(6).createCell(4);
                    XSSFCell countrTemp = newSheet.getRow(6).createCell(5);
                    XSSFCell nameTemp = newSheet.getRow(6).createCell(6);

                    XSSFCell nameList = row1.createCell(0);
                    newSheet.addMergedRegionUnsafe(new CellRangeAddress(0, 0, 0, 2));
                    nameList.setCellValue(sheetName);
                    nameList.setCellStyle(styleFont);

                    XSSFCell labor = row2.createCell(colLabor);
                    labor.setCellValue("Трудозатраты");
                    labor.setCellStyle(styleFont);

                    XSSFCell laborQuantry = row2.createCell(colLabor + 1);
                    laborQuantry.setCellValue(0);
                    laborQuantry.setCellStyle(center);

                    XSSFCell test = row3.createCell(colLabor);
                    test.setCellValue("Тест");
                    test.setCellStyle(styleFont);

                    XSSFCell testQuantry = row3.createCell(colLabor + 1);
                    testQuantry.setCellValue(0);
                    testQuantry.setCellStyle(center);

                    XSSFCell laborTotalName = row4.createCell(colLabor);
                    laborTotalName.setCellValue("Всего (ч/ч)");
                    laborTotalName.setCellStyle(styleFont);

                    XSSFCell laborTotal = row4.createCell(colLabor + 1);
                    laborTotal.setCellValue(laborQuantry.getNumericCellValue() + testQuantry.getNumericCellValue());
                    laborTotal.setCellStyle(center);

                    XSSFCell partNumer = row5.createCell(colPart);
                    partNumer.setCellValue("Парт номер");
                    partNumer.setCellStyle(styleFont);

                    XSSFCell quantity = row5.createCell(colQuantry);
                    quantity.setCellValue("Кол-во");
                    quantity.setCellStyle(styleFont);

                    XSSFCell partName = row5.createCell(colName);
                    partName.setCellValue("Название");
                    partName.setCellStyle(styleFont);

                    for (Row row : openSheet) {
                        Iterator<Cell> cellIterator = row.cellIterator();
                        while (cellIterator.hasNext()) {
                            Cell cell = cellIterator.next();
                            if (cell.toString().contains("ОЗ") && cell.getColumnIndex() == 7) {
                                int colGet = cell.getColumnIndex();
                                int rowGet = cell.getRowIndex();

                                XSSFCell partOpen = openSheet.getRow(rowGet).getCell(colGet - 6);

                                if (partOpen.toString().contains("CA")) {

                                    XSSFCell qtyOpen = openSheet.getRow(rowGet).getCell(colGet - 5);
                                    XSSFCell nameOpen = openSheet.getRow(rowGet).getCell(colGet - 3);

                                    boolean located = false;

                                    for (int k = 7; k <= newSheet.getLastRowNum(); k++) {

                                        XSSFCell newpart = newSheet.getRow(k).getCell(0);
                                        XSSFCell newcountr = newSheet.getRow(k).getCell(1);
                                        XSSFCell newname = newSheet.getRow(k).getCell(2);

                                        if (newpart.toString().equals(partOpen.toString())) {
                                            newcountr.setCellValue(newcountr.getNumericCellValue() + qtyOpen.getNumericCellValue());
                                            located = true;
                                        }
                                    }
                                    if (!located) {
                                        XSSFRow newRow = newSheet.createRow(rows);

                                        XSSFCell newPart = newRow.createCell(colPart);
                                        newPart.setCellValue(partOpen.toString());
                                        newPart.setCellStyle(borders);

                                        XSSFCell newCounts = newRow.createCell(colQuantry);
                                        newCounts.setCellValue(qtyOpen.getNumericCellValue());
                                        newCounts.setCellStyle(center);

                                        XSSFCell newName = newRow.createCell(colName);
                                        newName.setCellValue(nameOpen.toString());
                                        newName.setCellStyle(borders);

                                        rows++;
                                    }
                                }
                            }
                        }
                    }
                    boolean sorted = false;

                    while (!sorted) {
                        sorted = true;
                        for (int l = 7; l < newSheet.getLastRowNum(); l++) {

                            XSSFCell part = newSheet.getRow(l).getCell(0);
                            XSSFCell countr = newSheet.getRow(l).getCell(1);
                            XSSFCell name = newSheet.getRow(l).getCell(2);

                            XSSFCell part1 = newSheet.getRow(l + 1).getCell(0);
                            XSSFCell countr1 = newSheet.getRow(l + 1).getCell(1);
                            XSSFCell name1 = newSheet.getRow(l + 1).getCell(2);

                            if (part1.toString().compareTo(part.toString()) < 0) {

                                partTemp.setCellValue(part.toString());
                                countrTemp.setCellValue(countr.getNumericCellValue());
                                nameTemp.setCellValue(name.toString());

                                part.setCellValue(part1.toString());
                                countr.setCellValue(countr1.getNumericCellValue());
                                name.setCellValue(name1.toString());

                                part1.setCellValue(partTemp.toString());
                                countr1.setCellValue(countrTemp.getNumericCellValue());
                                name1.setCellValue(nameTemp.toString());

                                sorted = false;

                            }
                        }
                    }

                    if(!(partTemp == null)){
                        newSheet.getRow(6).removeCell(partTemp);
                        newSheet.getRow(6).removeCell(countrTemp);
                        newSheet.getRow(6).removeCell(nameTemp);
                    }
                }

                try (FileOutputStream uotFile = new FileOutputStream(fileName)) {
                    saveBook.write(uotFile);
                } catch (IOException e) {
                    exception = true;
                    JOptionPane.showMessageDialog(null, "Не удалось записать файл.");
                    return;
                }
            } catch (IOException e) {
                exception = true;
                JOptionPane.showMessageDialog(null, "Не удалось обработать " + file.getName());
                return;
            }

        }
        if (fileExists & !exception) JOptionPane.showMessageDialog(null, "Файл обновлён");
        if (!fileExists & !exception) JOptionPane.showMessageDialog(null, "Файл создан");
    }
}

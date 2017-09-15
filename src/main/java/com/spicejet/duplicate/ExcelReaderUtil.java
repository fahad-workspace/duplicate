package com.spicejet.duplicate;

import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

public class ExcelReaderUtil {

    private String getCellValue(Cell cell) {
        if (cell != null) {
            cell.setCellType(CellType.STRING);
            return cell.getStringCellValue();
        }
        return "";
    }

    public List<Book> readBooksFromExcelFile(String excelFilePath) throws IOException {
        List<Book> listBooks = new ArrayList<>();
        FileInputStream inputStream = new FileInputStream(new File(excelFilePath));
        Workbook workbook = new XSSFWorkbook(inputStream);
        Sheet firstSheet = workbook.getSheetAt(0);
        for (Row nextRow : firstSheet) {
            Iterator<Cell> cellIterator = nextRow.cellIterator();
            Book aBook = new Book();
            while (cellIterator.hasNext()) {
                Cell nextCell = cellIterator.next();
                int columnIndex = nextCell.getColumnIndex();
                switch (columnIndex) {
                    case 0:
                        aBook.setFirst(getCellValue(nextCell));
                        break;
                    case 1:
                        aBook.setSecond(getCellValue(nextCell));
                        break;
                    default:
                        // do nothing
                }
            }
            boolean flag = true;
            for (Book bBook : listBooks) {
                if (bBook.getFirst().equalsIgnoreCase(aBook.getSecond()) && bBook.getSecond().equalsIgnoreCase(aBook.getFirst())) {
                    System.out.println("Deleting: " + bBook.getFirst() + " and " + bBook.getSecond());
                    //  listBooks.remove(bBook);
                    flag = false;
                    break;
                }
            }
            if (flag) {
                listBooks.add(aBook);
            }
        }
        workbook.close();
        inputStream.close();
        return listBooks;
    }

    public void writeBooksFromExcelFile(String inputExcelFilePath, String outputExcelFilePath, List<Book> listBooks)
            throws IOException {
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            XSSFSheet sheet = workbook.createSheet(getSheetName(inputExcelFilePath));
            int rowNum = 0;
            System.out.println("Processing : " + inputExcelFilePath
                    .split(Constants.SEPERATOR)[inputExcelFilePath.split(Constants.SEPERATOR).length - 1]);
            for (Book book : listBooks) {
                if (rowNum == 0) {
                    rowNum++;
                    createMainHeader(sheet, workbook, book);
                    continue;
                }
                Row row = sheet.createRow(rowNum++);
                int colNum = 0;
                Cell cell = row.createCell(colNum++);
                cell.setCellValue(book.getFirst());
                cell = row.createCell(colNum);
                cell.setCellValue(book.getSecond());
            }
            FileOutputStream outputStream = new FileOutputStream(outputExcelFilePath);
            workbook.write(outputStream);
            workbook.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
        System.out.println("Processed : " + inputExcelFilePath
                .split(Constants.SEPERATOR)[inputExcelFilePath.split(Constants.SEPERATOR).length - 1]);
    }

    private String getSheetName(String excelFilePath) throws IOException {
        FileInputStream inputStream = new FileInputStream(new File(excelFilePath));
        Workbook workbook = new XSSFWorkbook(inputStream);
        Sheet firstSheet = workbook.getSheetAt(0);
        workbook.close();
        return firstSheet.getSheetName();
    }

    private void createMainHeader(XSSFSheet sheet, XSSFWorkbook workbook, Book book) {
        Row header = sheet.createRow(0);
        header.createCell(0).setCellValue(book.getFirst());
        header.createCell(1).setCellValue(book.getSecond());
        setHeaderStyle(workbook, header, sheet, IndexedColors.GREY_40_PERCENT.getIndex());
    }

    private void setHeaderStyle(XSSFWorkbook workbook, Row header, XSSFSheet sheet, short backgroundColor) {
        CellStyle style = workbook.createCellStyle();
        Font font = workbook.createFont();
        font.setBold(true);
        style.setFont(font);
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setFillForegroundColor(backgroundColor);
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setBorderBottom(BorderStyle.MEDIUM);
        style.setBorderTop(BorderStyle.MEDIUM);
        style.setBorderRight(BorderStyle.MEDIUM);
        style.setBorderLeft(BorderStyle.MEDIUM);
        style.setBottomBorderColor(HSSFColor.HSSFColorPredefined.DARK_RED.getIndex());
        style.setTopBorderColor(HSSFColor.HSSFColorPredefined.DARK_RED.getIndex());
        style.setRightBorderColor(HSSFColor.HSSFColorPredefined.DARK_RED.getIndex());
        style.setLeftBorderColor(HSSFColor.HSSFColorPredefined.DARK_RED.getIndex());
        for (int i = 0; i < header.getLastCellNum(); i++) {
            if (header.getCell(i) != null) {
                header.getCell(i).setCellStyle(style);
                sheet.autoSizeColumn(i);
            }
        }
    }
}

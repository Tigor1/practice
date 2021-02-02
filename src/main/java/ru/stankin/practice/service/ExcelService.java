package ru.stankin.practice.service;


import lombok.RequiredArgsConstructor;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellReference;
import org.springframework.stereotype.Service;
import ru.stankin.practice.entity.Person;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.*;

@Service
@RequiredArgsConstructor
public class ExcelService {
    private final PersonService personService;
    private int shiftColumn = 13;

    public void addHeader() throws IOException {
        Workbook headerWorkbook = WorkbookFactory.create(getClass().getClassLoader().getResourceAsStream("header.xls"));
        Workbook resultWorkBook = new HSSFWorkbook();

        Sheet headerSheet = headerWorkbook.getSheetAt(0);
        Sheet resultSheet = resultWorkBook.createSheet("Persons");

        Iterator<Row> headerRows = headerSheet.rowIterator();
        for (int i = 0; headerRows.hasNext(); i++) {
            Row headerRow = headerRows.next();
            Iterator<Cell> headerCells = headerRow.cellIterator();
            Row resultRow = resultSheet.createRow(headerRow.getRowNum());
//            resultRow.setRowStyle(headerRow.getRowStyle());

            while (headerCells.hasNext()) {
                Cell headerCell = headerCells.next();

                Cell resultCell = resultRow.createCell(headerCell.getColumnIndex());
                setCell(headerCell, resultCell);
//                resultCell.setCellStyle(headerCell.getCellStyle());
            }
            if (i == 13)
                break;
        }

        resultWorkBook.write(new FileOutputStream("result.xls"));
    }

    public void addPerson() throws IOException {
        Workbook resultWorkbook = WorkbookFactory.create(new FileInputStream("result.xls"));
        Sheet sheet = resultWorkbook.getSheetAt(0);


        List<Person> persons = personService.getPersons();
        for (int i = shiftColumn + 1; i < persons.size(); i++) {
            Row row = sheet.createRow(shiftColumn + i);



        }
    }

    public void setPersonOnRow(Row row, Person person, int numPerson) {
        CellReference cr1 = new CellReference("A" + (row.getRowNum() + 1));
        Cell numCell = row.createCell(cr1.getCol());
        numCell.setCellValue(numPerson);

        CellReference cr2 = new CellReference("B" + (row.getRowNum() + 1));
        Cell dotCell = row.createCell(cr2.getCol());
        dotCell.setCellValue(".");

        CellReference cr3 = new CellReference("C" + (row.getRowNum() + 1));
        Cell fioCell = row.createCell(cr3.getCol());
        fioCell.setCellValue(person.getSurname() + " " + person.getName() + " " + person.getSurname());

        CellReference cr4 = new CellReference("D" + (row.getRowNum() + 1));
        Cell registrNumCell = row.createCell(cr4.getCol());
        registrNumCell.setCellValue(person.getNumber());

        CellReference cr5 = new CellReference("F" + (row.getRowNum() + 1));
        Cell proffessionCell = row.createCell(cr5.getCol());
        proffessionCell.setCellValue(person.getProfession());

        CellReference cr11 = new CellReference("H" + (row.getRowNum() + 1));
        CellReference cr12 = new CellReference("V" + (row.getRowNum() + 1));

        for (int i = cr11.getCol(), j = 0; i <= cr12.getCol(); i++, j++) {
            Cell firstLineCell = row.createCell(cr11.getCol());
            firstLineCell.setCellValue(person.getTypeDays().get(j));
        }

        CellReference cr1HalfMonth = new CellReference("W" + (row.getRowNum() + 1));
        Cell halfMonthCell1 = row.createCell(cr1HalfMonth.getCol());
        halfMonthCell1.setCellValue(person.getTypeDays().stream()
                                        .reduce(0, (prev, cur) -> {
                                            if (cur.equals("A")) return prev + 1;
                                        }, (x, y)->x + y);


        CellReference cr13 = new CellReference("Z" + (row.getRowNum() + 1));
        CellReference cr14 = new CellReference("AO" + (row.getRowNum() + 1));

        for (int i = cr13.getCol(), j = 0; i <= cr14.getCol(); i++, j++) {
            firstLine.add(getCellValue(sheet.getRow(cr13.getRow()).getCell(cr13.getCol() + j)));
        }

        CellReference crTotalMonth = new CellReference("AP" + (row.getRowNum() + 1));
        firstLineTotal.add(getCellValue(sheet.getRow(crTotalMonth.getRow()).getCell(crTotalMonth.getCol())));


        //second line
        CellReference cr21 = new CellReference("H" + (row.getRowNum() + 2));
        CellReference cr22 = new CellReference("V" + (row.getRowNum() + 2));

        for (int i = cr21.getCol(), j = 0; i <= cr22.getCol(); i++, j++) {
            secondLine.add(getCellValue(sheet.getRow(cr21.getRow()).getCell(cr21.getCol() + j)));
        }

        CellReference crHalfMonthLine2 = new CellReference("W" + (row.getRowNum() + 2));
        secondLineTotal.add(getCellValue(sheet.getRow(crHalfMonthLine2.getRow()).getCell(crHalfMonthLine2.getCol())));

        CellReference cr23 = new CellReference("Z" + (row.getRowNum() + 2));
        CellReference cr24 = new CellReference("AO" + (row.getRowNum() + 2));

        for (int i = cr23.getCol(), j = 0; i <= cr24.getCol(); i++, j++) {
            firstLine.add(getCellValue(sheet.getRow(cr23.getRow()).getCell(cr23.getCol() + j)));
        }

        CellReference crTotalMonthLine2 = new CellReference("AP" + (row.getRowNum() + 2));
        secondLineTotal.add(getCellValue(sheet.getRow(crTotalMonthLine2.getRow()).getCell(crTotalMonthLine2.getCol())));

    }

    public void addFooter() throws IOException {
        Workbook headerWorkbook = WorkbookFactory.create(getClass().getClassLoader().getResourceAsStream("header.xls"));
        Workbook resultWorkBook = new HSSFWorkbook();

        Sheet headerSheet = headerWorkbook.getSheetAt(0);
        Sheet resultSheet = resultWorkBook.createSheet("Persons");

        Iterator<Row> headerRows = headerSheet.rowIterator();
        while (headerRows.hasNext()) {
            Row headerRow = headerRows.next();
            Iterator<Cell> headerCells = headerRow.cellIterator();
            Row resultRow = resultSheet.createRow(headerRow.getRowNum());
//            resultRow.setRowStyle(headerRow.getRowStyle());

            while (headerCells.hasNext()) {
                Cell headerCell = headerCells.next();

                Cell resultCell = resultRow.createCell(headerCell.getColumnIndex());
                setCell(headerCell, resultCell);
//                resultCell.setCellStyle(headerCell.getCellStyle());
            }
        }

        resultWorkBook.write(new FileOutputStream("result.xls"));
    }



    public void setCell(Cell from, Cell to) {
        if (from.getCellType() == CellType.STRING)
            to.setCellValue(from.getStringCellValue());
        else if (from.getCellType() == CellType.NUMERIC)
            to.setCellValue(from.getNumericCellValue());
        else if (from.getCellType() == CellType.BOOLEAN)
            to.setCellValue(from.getBooleanCellValue());
        else if (from.getCellType() == CellType.FORMULA)
            to.setCellValue(from.getCellFormula());
    }


    public List<Person> readPerson() throws IOException {

        Workbook workbook = WorkbookFactory.create(getClass().getClassLoader().getResourceAsStream("header.xls"));

        Sheet sheet = workbook.getSheetAt(0);
        List<Person> persons = new ArrayList<>();

        Iterator<Row> rows = sheet.rowIterator();
        Row currentRow = null;
        for (int i = 0; i < 15; i++) {
            currentRow = rows.next();
        }

        CellReference cr = new CellReference("C" + currentRow.getRowNum());
        String value = sheet.getRow(currentRow.getRowNum()).getCell(cr.getCol()).getStringCellValue();
        while (rows.hasNext()) {
            String[] fio = value.split(" ");
            Person.PersonBuilder newPerson = Person.builder();
            newPerson.name(fio[1]);
            newPerson.surname(fio[0]);
            newPerson.middlename(fio[2]);

            CellReference crNum = new CellReference("D" + currentRow.getRowNum());
            newPerson.number(sheet.getRow(currentRow.getRowNum()).getCell(crNum.getCol()).getStringCellValue());

            CellReference crProf = new CellReference("F" + currentRow.getRowNum());
            newPerson.profession(sheet.getRow(currentRow.getRowNum()).getCell(crProf.getCol()).getStringCellValue());


            Map<String, List<String>> map = readElementOfMonth(sheet, currentRow);

            System.out.println("----------------------------------------");
            System.out.println("row - " + currentRow.getRowNum());
            map.forEach((key, vl) -> {
                System.out.println(key + ": ");
                vl.forEach(System.out::println);
            });
            System.out.println("----------------------------------------");

            persons.add(newPerson.build());
            currentRow = rows.next();
            currentRow = rows.next();

            cr = new CellReference("C" + currentRow.getRowNum());
            if (sheet.getRow(currentRow.getRowNum()).getCell(cr.getCol()) == null)
                break;
            value = sheet.getRow(currentRow.getRowNum()).getCell(cr.getCol()).getStringCellValue();
        }
        return persons;
    }

    public Map<String, List<String>> readElementOfMonth(Sheet sheet, Row row) {
        List<String> firstLine = new ArrayList<>();
        List<String> secondLine = new ArrayList<>();
        List<String> firstLineTotal = new ArrayList<>();
        List<String> secondLineTotal = new ArrayList<>();

        //first line
        CellReference cr11 = new CellReference("H" + (row.getRowNum() + 1));
        CellReference cr12 = new CellReference("V" + (row.getRowNum() + 1));

        for (int i = cr11.getCol(), j = 0; i <= cr12.getCol(); i++, j++) {
            firstLine.add(getCellValue(sheet.getRow(cr11.getRow()).getCell(cr11.getCol() + j)));
        }

        CellReference crHalfMonth = new CellReference("W" + (row.getRowNum() + 1));
        firstLineTotal.add(getCellValue(sheet.getRow(crHalfMonth.getRow()).getCell(crHalfMonth.getCol())));

        CellReference cr13 = new CellReference("Z" + (row.getRowNum() + 1));
        CellReference cr14 = new CellReference("AO" + (row.getRowNum() + 1));

        for (int i = cr13.getCol(), j = 0; i <= cr14.getCol(); i++, j++) {
            firstLine.add(getCellValue(sheet.getRow(cr13.getRow()).getCell(cr13.getCol() + j)));
        }

        CellReference crTotalMonth = new CellReference("AP" + (row.getRowNum() + 1));
        firstLineTotal.add(getCellValue(sheet.getRow(crTotalMonth.getRow()).getCell(crTotalMonth.getCol())));


        //second line
        CellReference cr21 = new CellReference("H" + (row.getRowNum() + 2));
        CellReference cr22 = new CellReference("V" + (row.getRowNum() + 2));

        for (int i = cr21.getCol(), j = 0; i <= cr22.getCol(); i++, j++) {
            secondLine.add(getCellValue(sheet.getRow(cr21.getRow()).getCell(cr21.getCol() + j)));
        }

        CellReference crHalfMonthLine2 = new CellReference("W" + (row.getRowNum() + 2));
        secondLineTotal.add(getCellValue(sheet.getRow(crHalfMonthLine2.getRow()).getCell(crHalfMonthLine2.getCol())));

        CellReference cr23 = new CellReference("Z" + (row.getRowNum() + 2));
        CellReference cr24 = new CellReference("AO" + (row.getRowNum() + 2));

        for (int i = cr23.getCol(), j = 0; i <= cr24.getCol(); i++, j++) {
            firstLine.add(getCellValue(sheet.getRow(cr23.getRow()).getCell(cr23.getCol() + j)));
        }

        CellReference crTotalMonthLine2 = new CellReference("AP" + (row.getRowNum() + 2));
        secondLineTotal.add(getCellValue(sheet.getRow(crTotalMonthLine2.getRow()).getCell(crTotalMonthLine2.getCol())));


        Map<String, List<String>> result = new HashMap<>();
        result.put("firstLine", firstLine);
        result.put("secondLine", secondLine);
        result.put("firstLineTotal", firstLineTotal);
        result.put("secondLineTotal", secondLineTotal);
        return result;
    }

    public String getCellValue(Cell cell) {
        if (cell.getCellType() == CellType.STRING)
            return cell.getStringCellValue();
        else
            return String.valueOf(cell.getNumericCellValue());
    }
}

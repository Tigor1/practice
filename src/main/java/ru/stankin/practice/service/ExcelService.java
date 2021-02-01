package ru.stankin.practice.service;


import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellReference;
import org.springframework.stereotype.Service;
import ru.stankin.practice.entity.Person;

import java.io.IOException;
import java.util.*;

@Service
public class ExcelService {

    public List<Person> readPerson() throws IOException {

        Workbook workbook = WorkbookFactory.create(getClass().getClassLoader().getResourceAsStream("table.xls"));

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

//            System.out.println("----------------------------------------");
//            System.out.println("row - " + currentRow.getRowNum());
//            map.forEach((key, vl) -> {
//                System.out.println(key + ": ");
//                vl.forEach(System.out::println);
//            });
//            System.out.println("----------------------------------------");

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

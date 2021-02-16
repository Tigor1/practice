package ru.stankin.practice.service;

import lombok.RequiredArgsConstructor;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbookFactory;
import org.springframework.stereotype.Service;
import ru.stankin.practice.entity.Person;
import ru.stankin.practice.utils.Utils;

import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.charset.Charset;
import java.nio.charset.StandardCharsets;
import java.time.LocalDate;
import java.util.*;

@Service
@RequiredArgsConstructor
public class ExcelService {
    private final PersonService personService;
    private int shiftRow = 13;

    public void writeToExcel() throws IOException {
        Workbook resultWorkBook = new XSSFWorkbook();
        Sheet resultSheet = resultWorkBook.createSheet("Persons");

        addHeader(resultWorkBook, resultSheet);
        addPersons(resultWorkBook, resultSheet);
        addFooter(resultWorkBook, resultSheet);
        resultWorkBook.write( new FileOutputStream("result.xlsx"));
        resultWorkBook.close();
        shiftRow = 13;
    }

    public void addHeader(Workbook resultWorkBook, Sheet resultSheet) throws IOException {
        Workbook headerWorkbook = WorkbookFactory.create(getClass().getClassLoader().getResourceAsStream("header.xlsx"));
        Sheet headerSheet = headerWorkbook.getSheetAt(0);

        Iterator<Row> headerRows = headerSheet.rowIterator();
        for (int i = 0; headerRows.hasNext(); i++) {
            Row headerRow = headerRows.next();
            Iterator<Cell> headerCells = headerRow.cellIterator();
            Row resultRow = resultSheet.createRow(headerRow.getRowNum());
// resultRow.setRowStyle(headerRow.getRowStyle());

            while (headerCells.hasNext()) {
                Cell headerCell = headerCells.next();

                Cell resultCell = resultRow.createCell(headerCell.getColumnIndex());
// resultSheet.removeMergedRegion(headerCell.getColumnIndex());
                resultSheet.addMergedRegionUnsafe(headerSheet.getMergedRegion(headerCell.getColumnIndex()));
                setCell(headerCell, resultCell);
// copyCellStyle(resultWorkBook, headerCell, resultCell);
                setCellStyle(resultWorkBook, headerCell, resultCell);
                resultSheet.setColumnWidth(resultCell.getColumnIndex(), headerSheet.getColumnWidth(headerCell.getColumnIndex()));
                resultSheet.setDefaultRowHeight(headerSheet.getDefaultRowHeight());
            }
            if (i == 13)
                break;
        }
    }


    public void addPersons(Workbook resultWorkbook, Sheet sheet) throws IOException {
        List<Person> persons = new ArrayList<>(personService.getPersons());
        for (int i = shiftRow + 1, j = 0; j < persons.size(); i += 2, j++) {
            Row row1 = sheet.createRow(i);
            Row row2 = sheet.createRow(i + 1);
            setPersonOnRow(resultWorkbook, sheet, row1, row2, persons.get(j), j + 1);
        }
        shiftRow += (persons.size() * 2);
    }

    public void setPersonOnRow(Workbook resultWorkbook, Sheet sheet, Row row, Row row2, Person person, int numPerson) {

        CellReference cr1 = new CellReference("A" + (row.getRowNum() + 1));
        Cell numCell = row.createCell(cr1.getCol());
        sheet.addMergedRegionUnsafe(new CellRangeAddress(row.getRowNum(), row2.getRowNum(), cr1.getCol(), cr1.getCol()));
        numCell.setCellValue(numPerson);
        setStyleCellForPerson(resultWorkbook, numCell);

        CellReference cr2 = new CellReference("B" + (row.getRowNum() + 1));
        Cell dotCell = row.createCell(cr2.getCol());
        sheet.addMergedRegionUnsafe(new CellRangeAddress(row.getRowNum(), row2.getRowNum(), cr2.getCol(), cr2.getCol()));
        dotCell.setCellValue(".");
        setStyleCellForPerson(resultWorkbook, dotCell);

        CellReference cr3 = new CellReference("C" + (row.getRowNum() + 1));
        Cell fioCell = row.createCell(cr3.getCol());
        sheet.addMergedRegionUnsafe(new CellRangeAddress(row.getRowNum(), row2.getRowNum(), cr3.getCol(), cr3.getCol()));
        String fio = person.getSurname() + " " + person.getName() + " " + person.getMiddlename();
        fioCell.setCellValue(new String(fio.getBytes(), Charset.forName("WINDOWS-1251")));
        setStyleCellForPerson(resultWorkbook, fioCell);



        CellReference cr4 = new CellReference("D" + (row.getRowNum() + 1));
        Cell registrNumCell = row.createCell(cr4.getCol());
        sheet.addMergedRegionUnsafe(new CellRangeAddress(row.getRowNum(), row2.getRowNum(), cr4.getCol(), cr4.getCol() + 1));
        registrNumCell.setCellValue(person.getNumber());
        setStyleCellForPerson(resultWorkbook, registrNumCell);




        CellReference cr5 = new CellReference("F" + (row.getRowNum() + 1));
        Cell proffessionCell = row.createCell(cr5.getCol());
        sheet.addMergedRegionUnsafe(new CellRangeAddress(row.getRowNum(), row2.getRowNum(), cr5.getCol(), cr5.getCol() + 1));
        proffessionCell.setCellValue(person.getProfession());
        setStyleCellForPerson(resultWorkbook, proffessionCell);




        CellReference cr11 = new CellReference("H" + (row.getRowNum() + 1));
        CellReference cr12 = new CellReference("V" + (row.getRowNum() + 1));

        for (int i = cr11.getCol(), j = 0; i <= cr12.getCol(); i++, j++) {
            Cell firstLineCell = row.createCell(cr11.getCol() + j);
            firstLineCell.setCellValue(person.getTypeDays().get(j));
            setStyleCellForPerson(resultWorkbook, firstLineCell);
        }

        CellReference cr1HalfMonth = new CellReference("W" + (row.getRowNum() + 1));
        Cell halfMonthCell1 = row.createCell(cr1HalfMonth.getCol());
        int hmc = 0;
        for (int i = 0; i < 15; i++) {
            if (person.getTypeDays().get(i).equals("Ф")) hmc += 1;
        }
        sheet.addMergedRegionUnsafe(new CellRangeAddress(row.getRowNum(), row.getRowNum(), cr1HalfMonth.getCol(), cr1HalfMonth.getCol() + 2));
        halfMonthCell1.setCellValue(hmc);
        setStyleCellForPerson(resultWorkbook, halfMonthCell1);




        CellReference cr13 = new CellReference("Z" + (row.getRowNum() + 1));
        CellReference cr14 = new CellReference("AO" + (row.getRowNum() + 1));

        for (int i = cr13.getCol(), j = 15; i <= cr14.getCol(); i++, j++) {
            Cell firstLineCell = row.createCell(cr13.getCol() + j - 15);
            firstLineCell.setCellValue(person.getTypeDays().get(j));
            setStyleCellForPerson(resultWorkbook, firstLineCell);

        }

        CellReference cr1TotalMonth = new CellReference("AP" + (row.getRowNum() + 1));
        Cell totalMonthCell = row.createCell(cr1TotalMonth.getCol());
        for (int i = 15; i < LocalDate.now().lengthOfMonth(); i++) {
            if (person.getTypeDays().get(i).equals("Ф")) hmc += 1;
        }
        sheet.addMergedRegionUnsafe(new CellRangeAddress(row.getRowNum(), row.getRowNum(), cr1TotalMonth.getCol(), cr1TotalMonth.getCol() + 2));
        totalMonthCell.setCellValue(hmc);
        setStyleCellForPerson(resultWorkbook, totalMonthCell);


//second line
        CellReference cr21 = new CellReference("H" + (row2.getRowNum() + 1));
        CellReference cr22 = new CellReference("V" + (row2.getRowNum() + 1));

        for (int i = cr21.getCol(), j = 0; i <= cr22.getCol(); i++, j++) {
            Cell secondLineCell = row2.createCell(cr21.getCol() + j);
            secondLineCell.setCellValue(person.getAmountHoursInDay().get(j));
            setStyleCellForPerson(resultWorkbook, secondLineCell);

        }

        CellReference cr2HalfMonth = new CellReference("W" + (row2.getRowNum() + 1));
        Cell halfMonthCell2 = row2.createCell(cr2HalfMonth.getCol());
        Double ahid = 0D;
        for (int i = 0; i < 15; i++) {
            if (Utils.isNumber(person.getAmountHoursInDay().get(i)))
                ahid += Double.parseDouble(person.getAmountHoursInDay().get(i));
        }
        sheet.addMergedRegionUnsafe(new CellRangeAddress(row2.getRowNum(), row2.getRowNum(), cr2HalfMonth.getCol(), cr2HalfMonth.getCol() + 2));
        halfMonthCell2.setCellValue(ahid);
        setStyleCellForPerson(resultWorkbook, halfMonthCell2);



        CellReference cr23 = new CellReference("Z" + (row2.getRowNum() + 1));
        CellReference cr24 = new CellReference("AO" + (row2.getRowNum() + 1));

        for (int i = cr23.getCol(), j = 15; i <= cr24.getCol(); i++, j++) {
            Cell secondLineCell = row2.createCell(cr23.getCol() + (j - 15));
            secondLineCell.setCellValue(person.getAmountHoursInDay().get(j));
            setStyleCellForPerson(resultWorkbook, secondLineCell);
        }

        CellReference crTotalMonthLine2 = new CellReference("AP" + (row2.getRowNum() + 1));
        Cell totalMonthCell2 = row2.createCell(crTotalMonthLine2.getCol());
        for (int i = 15; i < person.getAmountHoursInDay().size(); i++) {
            if (Utils.isNumber(person.getAmountHoursInDay().get(i)))
                ahid += Double.parseDouble(person.getAmountHoursInDay().get(i));
        }
        sheet.addMergedRegionUnsafe(new CellRangeAddress(row2.getRowNum(), row2.getRowNum(), crTotalMonthLine2.getCol(), crTotalMonthLine2.getCol() + 2));
        totalMonthCell2.setCellValue(ahid);
        setStyleCellForPerson(resultWorkbook, totalMonthCell2);

    }

    public void addFooter(Workbook resultWorkBook, Sheet resultSheet) throws IOException {
        Workbook footerWorkbook = WorkbookFactory.create(getClass().getClassLoader().getResourceAsStream("footer.xlsx"));
        Sheet footerSheet = footerWorkbook.getSheetAt(0);

        Iterator<Row> headerRows = footerSheet.rowIterator();
        while (headerRows.hasNext()) {
            Row headerRow = headerRows.next();
            Iterator<Cell> headerCells = headerRow.cellIterator();
            Row resultRow = resultSheet.createRow(headerRow.getRowNum() + shiftRow + 1);
// resultRow.setRowStyle(headerRow.getRowStyle());

            while (headerCells.hasNext()) {
                Cell headerCell = headerCells.next();

                Cell resultCell = resultRow.createCell(headerCell.getColumnIndex());
                setCell(headerCell, resultCell);
// copyCellStyle(resultWorkBook, headerCell, resultCell);
// resultCell.setCellStyle(headerCell.getCellStyle());
                setCellStyle(resultWorkBook, headerCell, resultCell);
                resultSheet.setColumnWidth(resultCell.getColumnIndex(), footerSheet.getColumnWidth(headerCell.getColumnIndex()));
                resultSheet.setDefaultRowHeight(footerSheet.getDefaultRowHeight());
// resultSheet.addMergedRegion(footerSheet.getMergedRegion(headerCell.getColumnIndex()));

            }
        }
    }

    public void setStyleCellForPerson(Workbook resultWorkBook, Cell cell) {
        CellStyle newCellStyle = resultWorkBook.createCellStyle();
//        newCellStyle.setFillBackgroundColor(IndexedColors.GREY_50_PERCENT.getIndex());
//        newCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        newCellStyle.setBorderBottom(BorderStyle.THIN);
        newCellStyle.setBottomBorderColor(IndexedColors.BLACK.getIndex());
        newCellStyle.setBorderLeft(BorderStyle.THIN);
        newCellStyle.setLeftBorderColor(IndexedColors.BLACK.getIndex());
        newCellStyle.setBorderTop(BorderStyle.THIN);
        newCellStyle.setTopBorderColor(IndexedColors.BLACK.getIndex());
        newCellStyle.setBorderRight(BorderStyle.THIN);
        newCellStyle.setRightBorderColor(IndexedColors.BLACK.getIndex());
//        Font font = resultWorkBook.createFont();
//        font.setCharSet(FontCharset.RUSSIAN.getValue());
//        newCellStyle.setFont(font);
        newCellStyle.setAlignment(HorizontalAlignment.CENTER);
        newCellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        cell.setCellStyle(newCellStyle);
    }

    public void copyCellStyle(Workbook resultWorkBook, Cell oldCell, Cell newCell) {
        CellStyle newCellStyle = resultWorkBook.createCellStyle();
        newCellStyle.cloneStyleFrom(oldCell.getCellStyle());
        newCell.setCellStyle(newCellStyle);
    }

    private void setCellStyle(Workbook resultWorkbook, Cell origCell, Cell newCell) {
        CellStyle cellStyle = resultWorkbook.createCellStyle();
//cellStyle.setFillBackgroundColor(IndexedColors.AQUA.getIndex());
        cellStyle.setFillForegroundColor(origCell.getCellStyle().getFillForegroundColor());
        cellStyle.setBottomBorderColor(origCell.getCellStyle().getBottomBorderColor());
        cellStyle.setLeftBorderColor(origCell.getCellStyle().getLeftBorderColor());

        cellStyle.setLocked(origCell.getCellStyle().getLocked());
        cellStyle.setQuotePrefixed(origCell.getCellStyle().getQuotePrefixed());
        cellStyle.setRightBorderColor(origCell.getCellStyle().getRightBorderColor());
        cellStyle.setRotation(origCell.getCellStyle().getRotation());
        cellStyle.setShrinkToFit(origCell.getCellStyle().getShrinkToFit());
        cellStyle.setTopBorderColor(origCell.getCellStyle().getTopBorderColor());
        cellStyle.setVerticalAlignment(origCell.getCellStyle().getVerticalAlignment());
        cellStyle.setWrapText(origCell.getCellStyle().getWrapText());
        cellStyle.setHidden(origCell.getCellStyle().getHidden());
        cellStyle.setIndention(origCell.getCellStyle().getIndention());
        cellStyle.setAlignment(origCell.getCellStyle().getAlignment());
        cellStyle.setBorderBottom(origCell.getCellStyle().getBorderBottom());
        cellStyle.setBorderLeft(origCell.getCellStyle().getBorderLeft());
        cellStyle.setBorderRight(origCell.getCellStyle().getBorderRight());
        cellStyle.setBorderTop(origCell.getCellStyle().getBorderTop());
        cellStyle.setFillPattern(origCell.getCellStyle().getFillPattern());

        newCell.setCellStyle(cellStyle);
    }

// public void copyRowStyle(Workbook resultWorkBook, Row oldCell, Row newCell) {
// RowSty newCellStyle = resultWorkBook.createCellStyle();
// newCellStyle.cloneStyleFrom(oldCell.getCellStyle());
// newCell.setCellStyle(newCellStyle);
// }

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
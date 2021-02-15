package ru.stankin.practice.controller;


import dnl.utils.text.table.TextTable;
import javafx.collections.FXCollections;
import javafx.collections.ObservableList;
import javafx.fxml.FXML;
import javafx.fxml.FXMLLoader;
import javafx.scene.Parent;
import javafx.scene.Scene;
import javafx.scene.control.*;
import javafx.scene.control.Button;
import javafx.scene.control.TextField;
import javafx.scene.control.cell.PropertyValueFactory;
import javafx.scene.text.Font;
import javafx.stage.Stage;
import javafx.util.Callback;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Component;
import ru.stankin.practice.entity.Person;
import ru.stankin.practice.service.ExcelService;
import ru.stankin.practice.service.PersonService;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.PrintStream;
import java.net.URL;
import java.util.List;

@Component
public class AppController {
    @Autowired
    private ExcelService excelService;
    @Autowired
    private PersonService personService;

    @FXML
    public TextField name;
    @FXML
    public TextField surname;
    @FXML
    public TextField middlename;
    @FXML
    public TextField profession;
    @FXML
    public TextField number;
    @FXML
    public Button acceptBtn;
    @FXML
    public Button showBtn;
    @FXML
    public TableView<Person> personsTable;
    @FXML
    private TableColumn<String, Person> clmn1;
    @FXML
    private TableColumn<String, Person> clmn2;
    @FXML
    private TableColumn<String, Person> clmn3;
    @FXML
    private TableColumn<String, Person> clmn4;
    @FXML
    private TableColumn<String, Person> clmn5;
    @FXML
    private List<ChoiceBox> daysList;

    @FXML
    public void testExcelClick() throws IOException {
        clmn1.setCellValueFactory(new PropertyValueFactory<>("name"));
        clmn2.setCellValueFactory(new PropertyValueFactory<>("surname"));
        clmn3.setCellValueFactory(new PropertyValueFactory<>("middlename"));
        clmn4.setCellValueFactory(new PropertyValueFactory<>("number"));
        clmn5.setCellValueFactory(new PropertyValueFactory<>("profession"));

//        try {
//            personsTable.getItems().addAll(excelService.readPerson());
//        } catch (Exception ex) {
//            ex.printStackTrace();
//        }

        excelService.writeToExcel();
    }

//    @FXML
//    public void testExcelClick() {
//
//    }

    @FXML
    public void acceptAction() {
        Person person = Person.builder()
                .name(name.getText())
                .surname(surname.getText())
                .middlename(middlename.getText())
                .profession(profession.getText())
                .number(number.getText())
                .build();

        personService.save(person);
        showAction();
    }

    /**
     * Можно сделать вызов этого метода при старте приложения, тогда можно убрать кнопку, к которой он привязан.
     * Он при всех изменениях и так автоматически вызывается
     **/
    @FXML
    public void showAction() {
        clmn1.setCellValueFactory(new PropertyValueFactory<String, Person>("surname"));
        clmn2.setCellValueFactory(new PropertyValueFactory<String, Person>("name"));
        clmn3.setCellValueFactory(new PropertyValueFactory<String, Person>("middlename"));
        clmn4.setCellValueFactory(new PropertyValueFactory<String, Person>("number"));
        clmn5.setCellValueFactory(new PropertyValueFactory<String, Person>("profession"));

        List<Person> personsFromDb = personService.getPersons();

        System.out.println("Найдено " + personsFromDb.size() + " персон в БД");
        personsFromDb.forEach(System.out::println);

        ObservableList<Person> list = FXCollections.observableList(personsFromDb);
        personsTable.setItems(list);

        personsTable.setRowFactory(new Callback<TableView<Person>, TableRow<Person>>() {
            @Override
            public TableRow<Person> call(TableView<Person> param) {
                return new TableRow<Person>() {
                    @Override
                    protected void updateItem(Person item, boolean empty) {
                        super.updateItem(item, empty);
                        if (item != null) {
                            Tooltip tooltip = new Tooltip();
                            String result = getInformation(item);
                            tooltip.setText(result);
                            tooltip.setFont(new Font(12));
                            setTooltip(tooltip);
                        }
                    }
                };
            }
        });

//        /** По двойному клику по элементу таблицы будет вызываться метод descriptionAction();
//         * Сверху закомментировано, что при наведении на элемент таблицы там выводится "information"**/
//        personsTable.setRowFactory(param -> {
//            TableRow<Person> row = new TableRow<>();
//            row.setOnMouseClicked(event -> {
//                if (event.getClickCount() == 2 && (!row.isEmpty())) {
//                    Person rowData = row.getItem();
//                    descriptionAction(rowData);
//                }
//            });
//            return row;
//        });

        personsTable.refresh();
    }

    @FXML
    public void deleteAction() {
        ObservableList<TablePosition> list = personsTable.getSelectionModel().getSelectedCells();
        list.forEach(tablePosition -> {
            personService.delete(personsTable.getItems().get(tablePosition.getRow()).getId());
        });
        showAction();
    }

    public void descriptionAction(Person person) {

    }

    private String getInformation(Person person) {
        String[][] days = new String[3][person.getTypeDays().size()];
        for (int i = 0; i < person.getTypeDays().size(); i++) {
            days[0][i] = String.valueOf(i + 1);
        }
        for (int i = 0; i < person.getTypeDays().size(); i++) {
            days[1][i] = person.getTypeDays().get(i);
        }
        for (int i = 0; i < person.getAmountHoursInDay().size(); i++) {
            days[2][i] = person.getAmountHoursInDay().get(i);
        }
        TextTable textTable = new TextTable(person.getTypeDays().toArray(new String[0]), days);
        ByteArrayOutputStream byteArrayOutputStream = new ByteArrayOutputStream();
        PrintStream printStream = new PrintStream(byteArrayOutputStream);
        textTable.printTable(printStream, 0);
        return byteArrayOutputStream.toString();
    }

    @FXML
    private void test() {

    }
}

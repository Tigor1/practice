package ru.stankin.practice.controller;

import javafx.collections.FXCollections;
import javafx.collections.ObservableList;
import javafx.fxml.FXML;
import javafx.scene.control.*;
import javafx.scene.control.Button;
import javafx.scene.control.TextField;
import javafx.scene.control.cell.PropertyValueFactory;
import javafx.scene.paint.Color;
import javafx.scene.text.Text;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Component;
import ru.stankin.practice.entity.Person;
import ru.stankin.practice.service.ExcelService;
import ru.stankin.practice.service.PersonService;
import ru.stankin.practice.utils.Utils;

import java.io.IOException;
import java.util.*;
import java.util.List;
import java.util.stream.Collectors;

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
    public Button clearBtn;
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
    private List<Text> textList;


    public void init() {
        //        Видимость полей в зависимости от кол-ва дней в текущем месяце
        showAction();
    }

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
        Person person = personService.findByFio(surname.getText(), name.getText(), middlename.getText());
        if (person == null) {
            person = Person.builder()
                    .name(name.getText())
                    .surname(surname.getText())
                    .middlename(middlename.getText())
                    .profession(profession.getText())
                    .number(number.getText())
                    .build();
        }
        List<String> typeDays = new ArrayList<>();
        person.setTypeDays(typeDays);
        List<String> amountHoursInDays = new ArrayList<>();
        person.setAmountHoursInDay(amountHoursInDays);
        personService.save(person);
        showAction();
    }

    /**
     * Можно сделать вызов этого метода при старте приложения, тогда можно убрать кнопку, к которой он привязан.
     * Он при всех изменениях и так автоматически вызывается
     **/
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

        personsTable.setRowFactory(param -> {
            TableRow<Person> row = new TableRow<>();
            row.setOnMouseClicked(event -> {
                if (!row.isEmpty()) {
                    Person rowData = row.getItem();
                    descriptionAction(rowData);
                }
            });
            return row;
        });

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
        name.setText(person.getName());
        middlename.setText(person.getMiddlename());
        surname.setText(person.getSurname());
        profession.setText(person.getProfession());
        number.setText(person.getNumber());
    }

    @FXML
    private void clearAction() {
        personService.deleteAll();
        showAction();
    }

}

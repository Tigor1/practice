package ru.stankin.practice.controller;


import javafx.collections.FXCollections;
import javafx.collections.ObservableList;
import javafx.fxml.FXML;
import javafx.scene.control.*;
import javafx.scene.control.cell.PropertyValueFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Component;
import ru.stankin.practice.entity.Person;
import ru.stankin.practice.service.PersonService;

import java.util.List;

@Component
public class AppController {

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

    /** Можно сделать вызов этого метода при старте приложения, тогда можно убрать кнопку, к которой он привязан.
     * Он при всех изменениях и так автоматически вызывается**/
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
//
//        personsTable.setRowFactory(new Callback<TableView<Person>, TableRow<Person>>() {
//            @Override
//            public TableRow<Person> call(TableView<Person> param) {
//                return new TableRow<Person>() {
//                    @Override
//                    protected void updateItem(Person item, boolean empty) {
//                        super.updateItem(item, empty);
//                        if (item != null) {
//                            setTooltip(new Tooltip("information"));
//                        }
//                    }
//                };
//            }
//        });

        /** По двойному клику по элементу таблицы будет вызываться метод descriptionAction();
         * Сверху закомментировано, что при наведении на элемент таблицы там выводится "information"**/
        personsTable.setRowFactory(param -> {
            TableRow<Person> row = new TableRow<>();
            row.setOnMouseClicked(event -> {
                if (event.getClickCount() == 2 && (!row.isEmpty())) {
                    Person rowData = row.getItem();
                    descriptionAction(rowData);
                }
            });
            return row;
        });

//        personsTable.refresh();
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

}
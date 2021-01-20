package ru.stankin.practice.controller;


import javafx.fxml.FXML;
import javafx.scene.control.*;
import javafx.util.Callback;
import lombok.Data;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Component;
import ru.stankin.practice.entity.Person;
import ru.stankin.practice.entity.PersonRepository;
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
    public TextField middleName;
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
    public void acceptAction() {
        Person person = Person.builder()
                .name(name.getText())
                .surname(surname.getText())
                .middlename(middleName.getText())
                .profession(profession.getText())
                .number(number.getText())
                .build();

        personService.save(person);
    }

    @FXML
    public void showAction() {
        List<Person> personsFromDb = personService.getPersons();
        personsTable.setRowFactory(new Callback<TableView<Person>, TableRow<Person>>() {
            @Override
            public TableRow<Person> call(TableView<Person> param) {
                return new TableRow<Person>() {
                    @Override
                    protected void updateItem(Person item, boolean empty) {
                        super.updateItem(item, empty);
                        if (item != null) {
                            setTooltip(new Tooltip("information"));
                        }
                    }
                };
            }
        });
        personsFromDb.forEach(person -> personsTable.getItems().add(person));
    }

}

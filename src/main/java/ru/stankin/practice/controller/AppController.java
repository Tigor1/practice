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
    private List<ChoiceBox> daysList;
    @FXML
    private List<TextField> hoursList;
    @FXML
    private List<Text> textList;


    public void init() {
        //        Видимость полей в зависимости от кол-ва дней в текущем месяце
        int daysInCurrentMonth = Utils.getDaysInCurrentMonth();
        if (daysInCurrentMonth < 31) {
            for (int i = Utils.getDaysInCurrentMonth(); i < daysList.size(); i++) {
                daysList.get(i).setVisible(false);
                hoursList.get(i).setVisible(false);
            }
        } else {
            daysList.forEach(choiceBox -> choiceBox.setVisible(true));
            hoursList.forEach(choiceBox -> choiceBox.setVisible(true));
        }
        daysList = daysList.stream().filter(ChoiceBox::isVisible).collect(Collectors.toList());
        hoursList = hoursList.stream().filter(TextField::isVisible).collect(Collectors.toList());

        daysList.forEach(choiceBox -> {
            choiceBox.getItems().add("В");
            choiceBox.getItems().add("Н");
            choiceBox.getItems().add("Г");
            choiceBox.getItems().add("О");
            choiceBox.getItems().add("Б");
            choiceBox.getItems().add("Р");
            choiceBox.getItems().add("С");
            choiceBox.getItems().add("П");
            choiceBox.getItems().add("К");
            choiceBox.getItems().add("А");
            choiceBox.getItems().add("ВУ");
            choiceBox.getItems().add("ОУ");
            choiceBox.getItems().add("ЗН");
            choiceBox.getItems().add("ЗП");
            choiceBox.getItems().add("ЗС");
            choiceBox.getItems().add("РП");
            choiceBox.getItems().add("Ф");
            choiceBox.getItems().add("Я");
        });

        Calendar calendar = new GregorianCalendar();
        for (int i = 0; i < daysInCurrentMonth; i++) {
            calendar.set(Calendar.DAY_OF_MONTH, i + 1);
            textList.get(i).setText((i + 1) + " - " + calendar.getDisplayName(Calendar.DAY_OF_WEEK, Calendar.SHORT, Locale.getDefault()));
            /** if -> выходной, else -> рабочий**/
            if (calendar.get(Calendar.DAY_OF_WEEK) == Calendar.SATURDAY || calendar.get(Calendar.DAY_OF_WEEK) == Calendar.SUNDAY) {
                daysList.get(i).setValue("В");
                hoursList.get(i).setText("0");
                textList.get(i).setFill(Color.RED);
            } else {
                daysList.get(i).setValue("Р");
                hoursList.get(i).setText("8");
            }
        }
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
        daysList.forEach(choiceBox -> typeDays.add(choiceBox.getValue().toString()));
        person.setTypeDays(typeDays);
        List<String> amountHoursInDays = new ArrayList<>();
        hoursList.forEach(textField -> amountHoursInDays.add(textField.getText()));
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
        for (int i = 0; i < daysList.size() && i < person.getTypeDays().size() && i < person.getAmountHoursInDay().size(); i++) {
            daysList.get(i).setValue(person.getTypeDays().get(i));
            hoursList.get(i).setText(person.getAmountHoursInDay().get(i));
        }
    }

    @FXML
    private void clearAction() {
        personService.deleteAll();
        showAction();
    }

}

package ru.stankin.practice;

import javafx.application.Application;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import ru.stankin.practice.ui.AccountantApp;

@SpringBootApplication
public class PracticeApp {
    public static void main(String[] args) {
        Application.launch(AccountantApp.class, args);
    }
}

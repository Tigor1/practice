package ru.stankin.practice.ui;

import javafx.application.Application;
import javafx.application.Platform;
import javafx.fxml.FXMLLoader;
import javafx.scene.Parent;
import javafx.scene.Scene;
import javafx.stage.Stage;
import org.springframework.boot.builder.SpringApplicationBuilder;
import org.springframework.context.ApplicationEvent;
import org.springframework.context.ConfigurableApplicationContext;
import ru.stankin.practice.PracticeApp;
import ru.stankin.practice.controller.AppController;

import java.net.URL;

public class AccountantApp extends Application {
    private ConfigurableApplicationContext context;

    @Override
    public void init() throws Exception {
        context = new SpringApplicationBuilder(PracticeApp.class).run();
    }

    @Override
    public void stop() throws Exception {
        context.close();
        Platform.exit();
    }

    @Override
    public void start(Stage primaryStage) throws Exception {
        URL url = getClass().getClassLoader().getResource("view/window.fxml");
        System.out.println(url.getPath());

        FXMLLoader fxmlLoader = new FXMLLoader(url);
        fxmlLoader.setControllerFactory(context::getBean);

        Parent root = fxmlLoader.load();
        Scene scene = new Scene(root);

        primaryStage.setTitle("ЛошаченкоИД ИДБ-17-02");
        primaryStage.setScene(scene);
        context.publishEvent(new StageReadyEvent(primaryStage));
        primaryStage.show();
        context.getBean(AppController.class).init();
    }

    static class StageReadyEvent extends ApplicationEvent {

        public StageReadyEvent(Stage stage) {
            super(stage);
        }

        public Stage getStage() {
            return (Stage) getSource();
        }
    }
}

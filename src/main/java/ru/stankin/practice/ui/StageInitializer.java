package ru.stankin.practice.ui;

import org.springframework.context.ApplicationListener;
import org.springframework.stereotype.Component;

import static ru.stankin.practice.ui.AccountantApp.StageReadyEvent;

@Component
public class StageInitializer implements ApplicationListener<StageReadyEvent> {

    @Override
    public void onApplicationEvent(AccountantApp.StageReadyEvent event) {
        event.getStage();
    }
}

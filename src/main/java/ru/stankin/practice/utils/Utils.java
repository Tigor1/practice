package ru.stankin.practice.utils;

import java.util.Calendar;
import java.util.GregorianCalendar;
import java.util.Locale;

public class Utils {
    public static boolean isNumber(String str) {
        try {
            Double.parseDouble(str);
            return true;
        } catch (NumberFormatException ex) {
            return false;
        }
    }

    public static String getDayOfWeek(int dayInMonth) {
        Calendar calendar = new GregorianCalendar();
        calendar.set(Calendar.DAY_OF_WEEK, dayInMonth);
        return calendar.getDisplayName(Calendar.DAY_OF_WEEK, Calendar.SHORT, Locale.getDefault());
    }

    public static int getDaysInCurrentMonth() {
        Calendar calendar = new GregorianCalendar();
        return calendar.getActualMaximum(Calendar.DAY_OF_MONTH);
    }
}

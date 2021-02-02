package ru.stankin.practice.entity;

import lombok.*;


import javax.persistence.*;
import java.util.List;


@Entity
@Data
@Builder
@AllArgsConstructor
@NoArgsConstructor
@Table(name = "PERSON")
public class Person {
    @Id
    @GeneratedValue(strategy = GenerationType.AUTO)
    private Long id;

    private String name;
    private String surname;
    private String middlename;
    private String profession;
    private String number;

    private int workDays;
    private Double hoursDays;

    @ElementCollection
    @CollectionTable(
            name = "type_days",
            joinColumns = @JoinColumn(name = "person_id")
    )
    private List<String> typeDays;


    @ElementCollection
    @CollectionTable(
            name = "hours_days",
            joinColumns = @JoinColumn(name = "person_id")
    )
    private List<String> amountHoursInDay;
}

package ru.stankin.practice.entity;

import lombok.*;
import org.hibernate.annotations.LazyCollection;
import org.hibernate.annotations.LazyCollectionOption;


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
    @SequenceGenerator(name = "PERSON_SQ", sequenceName = "PERSON_SQ", allocationSize = 1)
    @GeneratedValue(strategy = GenerationType.SEQUENCE, generator = "PERSON_SQ")
    private Long id;

    private String name;
    private String surname;
    private String middlename;
    private String profession;
    private String number;


    private int workDays;
    private int halfWorkDays;
    private Double hoursDays;
    private Double halfHoursDays;

    @ElementCollection
    @LazyCollection(LazyCollectionOption.FALSE)
    @CollectionTable(
            name = "type_days",
            joinColumns = @JoinColumn(name = "person_id")
    )
    private List<String> typeDays;


    @ElementCollection
    @LazyCollection(LazyCollectionOption.FALSE)
    @CollectionTable(
            name = "hours_days",
            joinColumns = @JoinColumn(name = "person_id")
    )
    private List<String> amountHoursInDay;
}

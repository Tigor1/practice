package ru.stankin.practice.service;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import ru.stankin.practice.entity.Person;
import ru.stankin.practice.entity.PersonRepository;

import java.util.List;

@Service
public class PersonService {
    private final PersonRepository personRepository;

    @Autowired
    public PersonService(PersonRepository personRepository) {
        this.personRepository = personRepository;
    }

    public List<Person> getPersons() {
        return personRepository.findAll();
    }

    public Person save(Person person) {
        return personRepository.save(person);
    }

    public Person update(long id, Person personRequest) {
        return personRepository.findById(id).map(person-> {
            person.setName(personRequest.getName());
            person.setMiddlename(personRequest.getMiddlename());
            person.setProfession(personRequest.getProfession());
            person.setSurname(personRequest.getSurname());
            person.setNumber(personRequest.getNumber());
            return personRepository.save(person);
        }).orElseGet(()->personRepository.save(personRequest));
    }

    public void delete(long id) {
        personRepository.deleteById(id);
    }

    public Person findByFio(String surname, String name, String middleName) {
        return personRepository.findAll().stream()
                .filter(person ->
                        person.getSurname().equals(surname) &&
                        person.getName().equals(name) &&
                        person.getMiddlename().equals(middleName))
                .findFirst().orElse(null);
    }

    public void deleteAll() {
        personRepository.findAll().forEach(person -> delete(person.getId()));
    }
}

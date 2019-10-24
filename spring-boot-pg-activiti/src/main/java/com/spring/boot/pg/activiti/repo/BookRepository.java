package com.spring.boot.pg.activiti.repo;

import org.springframework.data.repository.CrudRepository;

import com.spring.boot.pg.activiti.model.Book;

public interface BookRepository extends CrudRepository<Book, Long> {

}

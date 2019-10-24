package com.spring.boot.pg.activiti.controller;

import java.util.ArrayList;
import java.util.List;
import java.util.Optional;

import javax.validation.Valid;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.CrossOrigin;
import org.springframework.web.bind.annotation.DeleteMapping;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PathVariable;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.PutMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import com.spring.boot.pg.activiti.model.Book;
import com.spring.boot.pg.activiti.repo.BookRepository;

@CrossOrigin(origins = "http://localhost:8081")
@RestController
@RequestMapping("/api")
public class BookController {

	@Autowired
	BookRepository bookRepository;

	@GetMapping("/books")
	public List<Book> getAllBooks() {
		System.out.println("Get all Books...");

		List<Book> list = new ArrayList<>();
		Iterable<Book> customers = bookRepository.findAll();

		customers.forEach(list::add);
		return list;
	}

	@PostMapping("/books/create")
	public Book createBook(@Valid @RequestBody Book book) {
		System.out.println("Create Book: " + book.getTitle() + "...");

		return bookRepository.save(book);
	}

	@GetMapping("/books/{id}")
	public ResponseEntity<Book> getBook(@PathVariable("id") Long id) {
		System.out.println("Get Book by id...");

		Book bookData = bookRepository.findOne(id);
		if (bookData != null) {
			return new ResponseEntity<>(bookData, HttpStatus.OK);
		} else {
			return new ResponseEntity<>(HttpStatus.NOT_FOUND);
		}
	}

	@PutMapping("/books/{id}")
	public ResponseEntity<Book> updateBook(@PathVariable("id") Long id, @RequestBody Book book) {
		System.out.println("Update Book with ID = " + id + "...");

		Book savedBook = bookRepository.findOne(id);
		if (savedBook != null) {
			//Book savedBook = bookData.get();
			savedBook.setTitle(book.getTitle());
			savedBook.setAuthor(book.getAuthor());
			savedBook.setDescription(book.getDescription());
			savedBook.setPublished(book.getPublished());

			Book updatedBook = bookRepository.save(savedBook);
			return new ResponseEntity<>(updatedBook, HttpStatus.OK);
		} else {
			return new ResponseEntity<>(HttpStatus.NOT_FOUND);
		}
	}

	@DeleteMapping("/books/{id}")
	public ResponseEntity<String> deleteBook(@PathVariable("id") Long id) {
		System.out.println("Delete Book with ID = " + id + "...");

		try {
			bookRepository.delete(id);
		} catch (Exception e) {
			return new ResponseEntity<>("Fail to delete!", HttpStatus.EXPECTATION_FAILED);
		}

		return new ResponseEntity<>("Book has been deleted!", HttpStatus.OK);
	}
}

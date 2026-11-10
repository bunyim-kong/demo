package com.example.demo.repository;

import org.springframework.data.repository.CrudRepository;
import com.example.demo.model.Category;

public interface CategoryRepository extends CrudRepository<Category, Long>{
    
}

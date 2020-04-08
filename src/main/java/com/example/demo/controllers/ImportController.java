package com.example.demo.controllers;

import com.example.demo.models.User;
import com.example.demo.repositories.UserRepository;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.ResponseBody;
import org.springframework.web.multipart.MultipartFile;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.nio.file.StandardCopyOption;

@Controller
public class ImportController {

    @Autowired
    private UserRepository userRepository;

    @PostMapping(value = "/import")
    @ResponseBody
    public void importUsers(@RequestParam("file")MultipartFile file) {
        String targetLocation = "/Users/Adri/workspace/file.xlsx";

        Path path = Paths.get(targetLocation).toAbsolutePath().normalize();

        try{
            Files.copy(file.getInputStream(), path, StandardCopyOption.REPLACE_EXISTING);
            FileInputStream fileStream = new FileInputStream(new File(targetLocation));
            Workbook workbook = new XSSFWorkbook(fileStream);

            Sheet sheet = workbook.getSheetAt(0);

            Integer index = 0;

            for(Row row : sheet){
                if(index != 0){
                    User user = new User();

                    user.setName(row.getCell(0).getStringCellValue());
                    user.setEmail(row.getCell(1).getStringCellValue());

                    userRepository.save(user);
                }

                index += 1;
            }
        }catch (IOException ex){
            ex.printStackTrace();
        }
    }

}

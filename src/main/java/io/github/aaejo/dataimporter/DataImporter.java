package io.github.aaejo.dataimporter;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.jdbc.core.JdbcTemplate;
import org.springframework.stereotype.Service;

import lombok.extern.slf4j.Slf4j;

import org.apache.poi.poifs.crypt.Decryptor;
import org.apache.poi.poifs.crypt.EncryptionInfo;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

@Slf4j
@Service
public class DataImporter {

    private final JdbcTemplate jdbcTemplate;

    public DataImporter(JdbcTemplate jdbcTemplate) {
        this.jdbcTemplate = jdbcTemplate;
    }

    private static final String FILE_NAME = "C:/Users/PC/Downloads/221020_DIA_All Users.xlsx";

    public void readFile() throws Exception {

        String personID, salutation, fname, mname, lname, address1, address2, address3, city, stateProv, postal, countryRegion, department, institution, institutionID, primeEmail, userID, ORCID, ORCIDVal, personAttribute, memberStatus;

        try {
            FileInputStream excelFile = new FileInputStream(new File(FILE_NAME));
            String password = "DIA_CUP";
            XSSFWorkbook workbook = (XSSFWorkbook) WorkbookFactory.create(excelFile, password);
            XSSFSheet sheet = workbook.getSheetAt(0);
            Iterator<Row> iterator = sheet.iterator();

            while (iterator.hasNext()) {
                Row currentRow = iterator.next();
                if (currentRow.getRowNum() == 0 || currentRow.getRowNum() == 1 || currentRow.getRowNum() == 2 || currentRow.getRowNum() == sheet.getLastRowNum()) {
                    continue;
                }
                Iterator<Cell> cellIterator = currentRow.iterator();
                List<String> cellValues = new ArrayList<String>();
                while (cellIterator.hasNext()) {
                    Cell currentCell = cellIterator.next();
                    cellValues.add(currentCell.toString());
                }
                
                personID = cellValues.get(0);
                salutation = cellValues.get(1);
                fname = cellValues.get(2);
                mname = cellValues.get(3);
                lname = cellValues.get(4);
                address1 = cellValues.get(5);
                address2 = cellValues.get(6);
                address3 = cellValues.get(7);
                city = cellValues.get(8);
                stateProv = cellValues.get(9);
                postal = cellValues.get(10);
                countryRegion = cellValues.get(11);
                department = cellValues.get(12);
                institution = cellValues.get(13);
                institutionID = cellValues.get(14);
                primeEmail = cellValues.get(15);
                userID = cellValues.get(16);
                ORCID = cellValues.get(17);
                ORCIDVal = cellValues.get(18);
                personAttribute = cellValues.get(19);
                memberStatus = cellValues.get(20);


                jdbcTemplate.update("INSERT INTO originalfile VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?);", personID, salutation, fname, mname, lname, address1, address2, address3, city, stateProv, postal, countryRegion, department, institution, institutionID, primeEmail, userID, ORCID, ORCIDVal, personAttribute, memberStatus);

            }
        }
        catch (Exception e) {
            throw new Exception("Error reading Excel file: " + e.getMessage());
        }
    }
}

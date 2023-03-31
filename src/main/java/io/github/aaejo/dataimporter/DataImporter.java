package io.github.aaejo.dataimporter;

import java.io.IOException;
import java.io.InputStream;
import java.util.Optional;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.dao.DataAccessException;
import org.springframework.jdbc.core.JdbcTemplate;
import org.springframework.stereotype.Service;

import lombok.extern.slf4j.Slf4j;

/**
 * @author Jeffery Kung
 * @author Omri Harary
 */
@Slf4j
@Service
public class DataImporter {

    private final JdbcTemplate jdbcTemplate;
    private int count;

    public DataImporter(JdbcTemplate jdbcTemplate) {
        this.jdbcTemplate = jdbcTemplate;
    }

    public void readFile(InputStream inStream, Optional<String> password)
            throws IOException, EncryptedDocumentException, DataAccessException {
        String personID, salutation, fname, mname, lname, address1, address2, address3, city, stateProv, postal,
                countryRegion, department, institution, institutionID, primeEmail, userID, ORCID, ORCIDVal,
                personAttribute, memberStatus;

        XSSFWorkbook workbook;
        if (password.isPresent()) {
            workbook = (XSSFWorkbook) WorkbookFactory.create(inStream, password.get());
        } else {
            workbook = (XSSFWorkbook) WorkbookFactory.create(inStream);
        }
        XSSFSheet sheet = workbook.getSheetAt(0);

        for (Row row : sheet) {
            log.debug("Processing row #{}", row.getRowNum());
            if (row.getRowNum() == 0 || row.getRowNum() == 1 || row.getRowNum() == 2
                    || row.getRowNum() == sheet.getLastRowNum()) {
                continue;
            }

            personID = row.getCell(0).toString();
            salutation = row.getCell(1).toString();
            fname = row.getCell(2).toString();
            mname = row.getCell(3).toString();
            lname = row.getCell(4).toString();
            address1 = row.getCell(5).toString();
            address2 = row.getCell(6).toString();
            address3 = row.getCell(7).toString();
            city = row.getCell(8).toString();
            stateProv = row.getCell(9).toString();
            postal = row.getCell(10).toString();
            countryRegion = row.getCell(11).toString();
            department = row.getCell(12).toString();
            institution = row.getCell(13).toString();
            institutionID = row.getCell(14).toString();
            primeEmail = row.getCell(15).toString();
            userID = row.getCell(16).toString();
            ORCID = row.getCell(17).toString();
            ORCIDVal = row.getCell(18).toString();
            personAttribute = row.getCell(19).toString();
            memberStatus = row.getCell(20).toString();

            jdbcTemplate.update(
                    "INSERT INTO originalfile VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?);",
                    personID, salutation, fname, mname, lname, address1, address2, address3, city, stateProv,
                    postal, countryRegion, department, institution, institutionID, primeEmail, userID, ORCID,
                    ORCIDVal, personAttribute, memberStatus);

            log.debug("Done row #{}", row.getRowNum());
            count++;
        }
    }

    public int getCount() {
        return count;
    }
}

package io.github.aaejo.dataimporter;

import java.io.IOException;
import java.io.InputStream;
import java.util.Optional;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.EncryptedDocumentException;
import org.springframework.context.annotation.Profile;
import org.springframework.dao.DataAccessException;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import lombok.extern.slf4j.Slf4j;

@Slf4j
@RestController
@Profile("default")
public class FileUploadController {

    private final DataImporter dataImporter;

    public FileUploadController(DataImporter dataImporter) {
        this.dataImporter = dataImporter;
    }

    @PostMapping("/")
    public void handleUpload(@RequestParam("file") MultipartFile file, @RequestParam("password") Optional<String> password)
            throws Exception {
        log.info("Received file {} {}", file.getOriginalFilename(),
                (password.isPresent() ? "with password" : "without password"));
        try (InputStream fileIn = file.getInputStream()) {
            dataImporter.readFile(fileIn, password);
        } catch (EncryptedDocumentException | DataAccessException | IOException e) {
            log.error("Encountered an error in processing uploaded file.", e);
            throw e;
        } finally {
            log.info("{} rows loaded from file", dataImporter.getCount());
        }
    }
}

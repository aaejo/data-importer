package io.github.aaejo.dataimporter;

import java.io.File;
import java.io.FileInputStream;
import java.util.Optional;

import org.springframework.beans.factory.annotation.Value;
import org.springframework.boot.ApplicationRunner;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.context.annotation.Bean;
import org.springframework.context.annotation.Profile;

@SpringBootApplication
public class DataImporterApplication {

	public static void main(String[] args) {
		SpringApplication.run(DataImporterApplication.class, args);
	}

	@Bean
	@Profile("console")
	public ApplicationRunner runner(DataImporter dataImporter,
			@Value("${aaejo.jds.data-importer.file}") String file,
			@Value("${aaejo.jds.data-importer.password}") Optional<String> password) {
		return args -> {
			FileInputStream excelFile = new FileInputStream(new File(file));
			dataImporter.readFile(excelFile, password);
		};
	}
}

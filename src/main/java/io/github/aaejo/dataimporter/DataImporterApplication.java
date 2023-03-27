package io.github.aaejo.dataimporter;

import org.springframework.boot.ApplicationRunner;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.context.annotation.Bean;


@SpringBootApplication
public class DataImporterApplication {

	private final DataImporter dataImporter;

	public DataImporterApplication(DataImporter dataImporter) {
		this.dataImporter = dataImporter;
	}
	public static void main(String[] args) {
		SpringApplication.run(DataImporterApplication.class, args);
	}

	@Bean
    public ApplicationRunner runner() {
        return args -> {
			dataImporter.readFile();
        };
    }

}

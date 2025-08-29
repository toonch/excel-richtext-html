package com.yipt.outbound;

import org.springframework.boot.CommandLineRunner;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.context.annotation.Bean;
import java.util.Collections;
import com.yipt.outbound.services.MigrationSCBService;
import com.yipt.outbound.services.MigrationService;


@SpringBootApplication
public class ArcherOutboundApplication {

	public static void main(String[] args) {
		SpringApplication app = new SpringApplication(ArcherOutboundApplication.class);
		// Set port from environment variable (PORT) or default 8080
        app.setDefaultProperties(
            Collections.singletonMap("server.port", System.getenv().getOrDefault("PORT", "8080"))
        );
        app.run(args);
	}
	
	
//	@Bean
//    public CommandLineRunner soapServiceRunner(MigrationSCBService migrationService) {
//        return args -> {
//        	migrationService.run();
//        };
//    }
	
}

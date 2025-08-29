package com.yipt.outbound;


import static org.junit.Assert.assertNotNull;

import org.junit.Test;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.test.context.SpringBootTest;


@SpringBootTest
class ArcherOutboundApplicationTests {
	
	@Autowired
    private ArcherOutboundApplication application;

	@Test
	void contextLoads() {
		assertNotNull(application);
	}

}

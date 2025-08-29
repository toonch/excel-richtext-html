package com.yipt.outbound.response;

import java.util.Map;

import lombok.Data;

@Data
public class Field {
	private Map<String, String> attributes;
    private String value;
}

package com.yipt.outbound.request;

import lombok.Data;

@Data
public class LoginRequest {
	private String instanceName;
    private String username;
    private String userDomain;
    private String password;
}

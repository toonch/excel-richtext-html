package com.yipt.outbound.response;

import com.fasterxml.jackson.annotation.JsonProperty;

import lombok.Data;

@Data
public class LoginResponse {
	@JsonProperty("Links")
    private Object[] links;

    @JsonProperty("RequestedObject")
    private RequestedObject requestedObject;

    @JsonProperty("IsSuccessful")
    private boolean isSuccessful;

    @JsonProperty("ValidationMessages")
    private Object[] validationMessages;
}

package com.yipt.outbound.response;

import com.fasterxml.jackson.annotation.JsonProperty;

import lombok.Data;

@Data
public class RequestedObject {
	@JsonProperty("SessionToken")
    private String sessionToken;

    @JsonProperty("InstanceName")
    private String instanceName;

    @JsonProperty("UserId")
    private int userId;

    @JsonProperty("ContextType")
    private int contextType;

    @JsonProperty("UserConfig")
    private UserConfig userConfig;

    @JsonProperty("Translate")
    private boolean translate;

    @JsonProperty("RequestLanguageId")
    private Object requestLanguageId;
}

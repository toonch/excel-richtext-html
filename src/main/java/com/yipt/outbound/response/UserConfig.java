package com.yipt.outbound.response;

import com.fasterxml.jackson.annotation.JsonProperty;

import lombok.Data;

@Data
public class UserConfig {
	@JsonProperty("TimeZoneId")
    private String timeZoneId;

    @JsonProperty("TimeZoneIdSource")
    private int timeZoneIdSource;

    @JsonProperty("LocaleId")
    private String localeId;

    @JsonProperty("LocaleIdSource")
    private int localeIdSource;

    @JsonProperty("LanguageId")
    private int languageId;

    @JsonProperty("DefaultHomeDashboardId")
    private int defaultHomeDashboardId;

    @JsonProperty("DefaultHomeWorkspaceId")
    private int defaultHomeWorkspaceId;

    @JsonProperty("LanguageIdSource")
    private int languageIdSource;

    @JsonProperty("PlatformLanguageId")
    private int platformLanguageId;

    @JsonProperty("PlatformLanguagePath")
    private String platformLanguagePath;

    @JsonProperty("PlatformLanguageIdSource")
    private int platformLanguageIdSource;
}

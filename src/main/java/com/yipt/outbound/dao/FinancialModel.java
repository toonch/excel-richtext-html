package com.yipt.outbound.dao;

import lombok.Data;

@Data
public class FinancialModel {
	private String trackingId;
    private String opLossReferenceId;
    private String financialStatus;
    private String effectNature;
    private String effectLossAmount;
    private String effectGainAmount;
    private String bookingDate;
    private String ledgerCode;
    private String otherLedgerCode;
    private String reference;
    private String bookingUnitCode;
    private String bookingUnitName;
    private String businessLineCodeBusinessLine;
}

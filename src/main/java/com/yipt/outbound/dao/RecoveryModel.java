package com.yipt.outbound.dao;

import lombok.Data;

@Data
public class RecoveryModel {
	private String recoveryTrackingId;
    private String recoveryType;
    private String recoveryAmount;
    private String recoveryDescription;
    private String bookingUnitCode;
    private String bookingUnitName;
    private String ledgerCode;
    private String otherLedgerCode;
    private String bookingDate;
    private String opLossReferenceId;
}

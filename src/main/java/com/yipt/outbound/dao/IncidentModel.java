package com.yipt.outbound.dao;

import lombok.Data;

@Data
public class IncidentModel {
	private String referenceId;                          // Reference ID
    private String incidentTitle;                       // Incident Title
    private String discoveredBy;                        // Discovered By
    private String description;                         // Description
    private String discoveryDetail;                     // Discovery Detail
    private String moLevel2;                            // MO Level2 (MO Level 1 and 2 Code)
    private String moLevel3;                            // MO Level3 (MO Level3 Code)
    private String moLevel4;                            // MO Level4 (MO Level4 Code)
    private String moLevel5;                            // MO Level5 (MO Level5 Code)
    private String moLevel6;                            // MO Level6 (MO Level6 Code)
    private String firstDateOfEvent;                    // First Date of Event
    private String dateOfDiscovery;                     // Date of Discovery
    private String riskEventType;                       // Risk Event Type
    private String amount;                              // Amount
    private String lastDateOfEvent;                     // Last Date of Event
    private String closedDate;                          // Closed Date
    private String sourceOfIncident;                    // Source of Incident
    private String contactDetail;                       // Contact Detail
    private String incidentDetailsComments;             // Incident Details Comments
    private String customerComplaint;                   // Customer Complaint?
    private String hasLegalImpact;                      // Has Legal Impact?
    private String complaintId;                         // Complaint ID
    private String legalCaseId;                         // Legal Case ID
    private String legalCaseStatus;                     // Legal Case Status
    private String prosecutionResults;                  // Prosecution Results
    private String daysForSla1;                         // Days for SLA1
    private String daysForSla2;                         // Days for SLA2
    private String riskCategories;                      // Risk Categories (Risk Code)
    private String creditRelatedOrMarketRelated;        // Credit Related or Market Related
    private String process;                             // Process (Process Code)
    private String product;                             // Product (Product/Services Code)
    private String cause1;                              // Cause1 (Cause Code)
    private String cause2;                              // Cause2 (Cause Code)
    private String causeDescription;                    // Cause Description
    private String failedControl;                       // Failed Control (Control Code)
    private String controlDetails;                      // Control Details
    private String correctiveAction;                    // Corrective Action
    private String relatedOpLossIncident;               // Related OpLoss Incident (Reference ID)
    private String impactType;                          // Impact Type
    private String brandImpactClassification;           // Brand Impact Classification
    private String employeeImpactClassification;        // Employee Impact Classification
    private String environmentalImpactClassification;   // Environmental Impact Classification
    private String legalImpactClassification;           // Legal Impact Classification
    private String regulatoryImpactClassification;      // Regulatory Impact Classification
    private String reputationalImpactClassification;    // Reputational Impact Classification
    private String stakeholderImpactClassification;     // Stakeholder Impact Classification
    private String brandImpactDescription;              // Brand Impact Description
    private String employeeImpactDescription;           // Employee Impact Description
    private String environmentalImpactDescription;      // Environmental Impact Description
    private String legalImpactDescription;              // Legal Impact Description
    private String regulatoryImpactDescription;         // Regulatory Impact Description
    private String reputationalImpactDescription;       // Reputational Impact Description
    private String stakeholderImpactDescription;        // Stakeholder Impact Description
    private String issues;                              // Issues
    private String referenceIdItsmBulkUpload;           // Reference ID (ITSM Interface / Bulk Upload)
    private String bulkUploadReferenceId;               // Bulk upload reference ID
}

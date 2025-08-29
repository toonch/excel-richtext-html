package com.yipt.outbound.services;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.*;
import org.springframework.stereotype.Component;

import com.opencsv.CSVWriter;
import com.yipt.outbound.dao.ExcelModel;
import com.yipt.outbound.dao.FinancialModel;
import com.yipt.outbound.dao.IncidentModel;
import com.yipt.outbound.dao.RecoveryModel;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStreamWriter;
import java.io.Writer;
import java.math.BigDecimal;
import java.math.RoundingMode;
import java.nio.charset.StandardCharsets;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.stream.Collectors;

@Component
public class MigrationService {
	public void run() {
		String filePath = "C:/Users/Toonch/Downloads/migrate_data.xlsx";
		List<ExcelModel> excelData = readExcelToModel(filePath);
		List<IncidentModel> incidentModels = new ArrayList<>();
		List<FinancialModel> financialModels = new ArrayList<>();
		List<RecoveryModel> recoveryModels = new ArrayList<>();
		StringBuilder relatedOpLossIncident = new StringBuilder();
		String tempNo = "";
		for (int i = 0; i < excelData.size(); i++) {
			if (excelData.get(i).getBColumn().equals("CSho_0500000023605")) {
				System.out.println("");
			}
			boolean changeSequence = !tempNo.equals(excelData.get(i).getBbColumn());
			if (changeSequence) {
				tempNo = excelData.get(i).getBbColumn();
			}

			// Check the next element only if it exists (i+1 is within bounds)
			if (i < excelData.size() - 1) {
				if (!changeSequence && (!excelData.get(i).getBColumn().equals(excelData.get(i + 1).getBColumn()))) {
					if (!excelData.get(i).getAsColumn().equals("1.0")) {
						IncidentModel incidentRow = new IncidentModel();
						incidentRow.setReferenceId(excelData.get(i).getBColumn());
						if (StringUtils.isEmpty(excelData.get(i).getOColumn())) { // if กิจกรรมความเสียหาย เป็นค่าว่าง
																					// then Loss Profile
							incidentRow.setIncidentTitle(excelData.get(i).getAxColumn());
						} else {
							incidentRow.setIncidentTitle(excelData.get(i).getOColumn());
						}
						incidentRow.setDiscoveredBy("อื่น ๆ (Other)");
						incidentRow.setDescription(
								excelData.get(i).getAbColumn().replace("\n", "<br/>").replace("\r", "<br/>"));
						incidentRow.setDiscoveryDetail("");
						incidentRow.setMoLevel2(excelData.get(i).getBcColumn());
						incidentRow.setMoLevel3(excelData.get(i).getBdColumn());
						incidentRow.setMoLevel4(excelData.get(i).getBeColumn());
						incidentRow.setMoLevel5(excelData.get(i).getBfColumn());
						incidentRow.setMoLevel6(excelData.get(i).getBgColumn());
						incidentRow.setFirstDateOfEvent(excelData.get(i).getTColumn());
						incidentRow.setRiskEventType(excelData.get(i).getBhColumn());
						incidentRow.setAmount("0");
						incidentRow.setDateOfDiscovery(excelData.get(i).getVColumn());
						incidentRow.setLastDateOfEvent(excelData.get(i).getUColumn());
						incidentRow.setClosedDate(excelData.get(i).getAjColumn());
						incidentRow.setSourceOfIncident("OpLoss Incident");
						incidentRow.setContactDetail("");
						incidentRow.setIncidentDetailsComments("");
						incidentRow.setCustomerComplaint("N");
						incidentRow.setHasLegalImpact("N");
						incidentRow.setComplaintId("");
						incidentRow.setLegalCaseId("");
						incidentRow.setLegalCaseStatus("");
						incidentRow.setRiskCategories(excelData.get(i).getBiColumn());
						if (StringUtils.isEmpty(excelData.get(i).getAnColumn())) {
							incidentRow.setCreditRelatedOrMarketRelated("None");
						} else {
							incidentRow.setCreditRelatedOrMarketRelated(excelData.get(i).getAnColumn());
						}
						incidentRow.setProcess(excelData.get(i).getBjColumn());
						incidentRow.setProduct(excelData.get(i).getBkColumn());
						incidentRow.setCause1(excelData.get(i).getBlColumn());
						incidentRow.setCauseDescription(
								excelData.get(i).getAcColumn().replace("\n", "<br/>").replace("\r", "<br/>"));
						incidentRow.setControlDetails(
								excelData.get(i).getArColumn().replace("\n", "<br/>").replace("\r", "<br/>"));
						if (StringUtils.isEmpty(incidentRow.getControlDetails())) {
							incidentRow.setCorrectiveAction("-");
						} else {
							incidentRow.setCorrectiveAction(incidentRow.getControlDetails());
						}
						incidentModels.add(incidentRow);
						relatedOpLossIncident.append(excelData.get(i).getBColumn()).append(";");

						FinancialModel financialRow = new FinancialModel();
						financialRow.setOpLossReferenceId(excelData.get(i).getBColumn());
						financialRow.setFinancialStatus(excelData.get(i).getBmColumn());
						BigDecimal GROSS_LOSS_AMT = !StringUtils.isEmpty(excelData.get(i).getAeColumn())
							    ? new BigDecimal(excelData.get(i).getAeColumn())  // Use BigDecimal for precision
							    : BigDecimal.ZERO;
						if (GROSS_LOSS_AMT.compareTo(BigDecimal.ZERO) < 0) { // If less than 0
						    financialRow.setEffectNature("Gain");
						    financialRow.setEffectLossAmount("");
						    financialRow.setEffectGainAmount(GROSS_LOSS_AMT.toPlainString()); // Avoid scientific notation
						} else {
						    financialRow.setEffectNature("Loss");
						    financialRow.setEffectLossAmount(GROSS_LOSS_AMT.toPlainString()); // Avoid scientific notation
						    financialRow.setEffectGainAmount("");
						}
						financialRow.setBookingDate(excelData.get(i).getAfColumn());
						if (!StringUtils.isEmpty(excelData.get(i).getAfColumn())
								&& excelData.get(i).getAaColumn().equals("ยังไม่บันทึกบัญชี หรือไม่สามารถระบุได้")) {
							financialRow.setLedgerCode("ยังไม่บันทึกบัญชี");
						} else {
							financialRow.setLedgerCode(excelData.get(i).getAaColumn());
						}
						financialRow.setBookingUnitCode(excelData.get(i).getCColumn());
						financialRow.setBookingUnitName(excelData.get(i).getKColumn());
						financialRow.setBusinessLineCodeBusinessLine(excelData.get(i).getLColumn());
						financialModels.add(financialRow);

					} else {
						IncidentModel incidentRow = new IncidentModel();
						incidentRow.setReferenceId(excelData.get(i).getBColumn());
						if (StringUtils.isEmpty(excelData.get(i).getOColumn())) { // if กิจกรรมความเสียหาย เป็นค่าว่าง
																					// then Loss Profile
							incidentRow.setIncidentTitle(excelData.get(i).getAxColumn());
						} else {
							incidentRow.setIncidentTitle(excelData.get(i).getOColumn());
						}
						incidentRow.setDiscoveredBy("อื่น ๆ (Other)");
						incidentRow.setDescription(
								excelData.get(i).getAbColumn().replace("\n", "<br/>").replace("\r", "<br/>"));
						incidentRow.setDiscoveryDetail("");
						incidentRow.setMoLevel2(excelData.get(i).getBcColumn());
						incidentRow.setMoLevel3(excelData.get(i).getBdColumn());
						incidentRow.setMoLevel4(excelData.get(i).getBeColumn());
						incidentRow.setMoLevel5(excelData.get(i).getBfColumn());
						incidentRow.setMoLevel6(excelData.get(i).getBgColumn());
						incidentRow.setFirstDateOfEvent(excelData.get(i).getTColumn());
						incidentRow.setRiskEventType(excelData.get(i).getBhColumn());
						incidentRow.setAmount("0");
						incidentRow.setDateOfDiscovery(excelData.get(i).getVColumn());
						incidentRow.setLastDateOfEvent(excelData.get(i).getUColumn());
						incidentRow.setClosedDate(excelData.get(i).getAjColumn());
						incidentRow.setSourceOfIncident("OpLoss Incident");
						incidentRow.setContactDetail("");
						incidentRow.setIncidentDetailsComments("");
						incidentRow.setCustomerComplaint("N");
						incidentRow.setHasLegalImpact("N");
						incidentRow.setComplaintId("");
						incidentRow.setLegalCaseId("");
						incidentRow.setLegalCaseStatus("");
						incidentRow.setRiskCategories(excelData.get(i).getBiColumn());
						if (StringUtils.isEmpty(excelData.get(i).getAnColumn())) {
							incidentRow.setCreditRelatedOrMarketRelated("None");
						} else {
							incidentRow.setCreditRelatedOrMarketRelated(excelData.get(i).getAnColumn());
						}
						incidentRow.setProcess(excelData.get(i).getBjColumn());
						incidentRow.setProduct(excelData.get(i).getBkColumn());
						incidentRow.setCause1(excelData.get(i).getBlColumn());
						incidentRow.setCauseDescription(
								excelData.get(i).getAcColumn().replace("\n", "<br/>").replace("\r", "<br/>"));
						incidentRow.setControlDetails(
								excelData.get(i).getArColumn().replace("\n", "<br/>").replace("\r", "<br/>"));
						if (StringUtils.isEmpty(incidentRow.getControlDetails())) {
							incidentRow.setCorrectiveAction("-");
						} else {
							incidentRow.setCorrectiveAction(incidentRow.getControlDetails());
						}
						// Remove the last semicolon if it exists
						if (relatedOpLossIncident.length() > 0
								&& relatedOpLossIncident.charAt(relatedOpLossIncident.length() - 1) == ';') {
							relatedOpLossIncident.setLength(relatedOpLossIncident.length() - 1); // Remove last
																									// semicolon
						}

						// Set the relatedOpLossIncident
						incidentRow.setRelatedOpLossIncident(relatedOpLossIncident.toString());
						incidentModels.add(incidentRow);
						relatedOpLossIncident = new StringBuilder();

						FinancialModel financialRow = new FinancialModel();
						financialRow.setOpLossReferenceId(excelData.get(i).getBColumn());
						financialRow.setFinancialStatus(excelData.get(i).getBmColumn());
						BigDecimal GROSS_LOSS_AMT = !StringUtils.isEmpty(excelData.get(i).getAeColumn())
							    ? new BigDecimal(excelData.get(i).getAeColumn())  // Use BigDecimal for precision
							    : BigDecimal.ZERO;
						if (GROSS_LOSS_AMT.compareTo(BigDecimal.ZERO) < 0) { // If less than 0
						    financialRow.setEffectNature("Gain");
						    financialRow.setEffectLossAmount("");
						    financialRow.setEffectGainAmount(GROSS_LOSS_AMT.toPlainString()); // Avoid scientific notation
						} else {
						    financialRow.setEffectNature("Loss");
						    financialRow.setEffectLossAmount(GROSS_LOSS_AMT.toPlainString()); // Avoid scientific notation
						    financialRow.setEffectGainAmount("");
						}
						financialRow.setBookingDate(excelData.get(i).getAfColumn());
						if (!StringUtils.isEmpty(excelData.get(i).getAfColumn())
								&& excelData.get(i).getAaColumn().equals("ยังไม่บันทึกบัญชี หรือไม่สามารถระบุได้")) {
							financialRow.setLedgerCode("ยังไม่บันทึกบัญชี");
						} else {
							financialRow.setLedgerCode(excelData.get(i).getAaColumn());
						}
						financialRow.setBookingUnitCode(excelData.get(i).getCColumn());
						financialRow.setBookingUnitName(excelData.get(i).getKColumn());
						financialRow.setBusinessLineCodeBusinessLine(excelData.get(i).getLColumn());
						financialModels.add(financialRow);
					}
				} else if (changeSequence && !excelData.get(i).getBColumn().equals(excelData.get(i + 1).getBColumn())) {
					if (!excelData.get(i).getAsColumn().equals("1.0")) {
						IncidentModel incidentRow = new IncidentModel();
						incidentRow.setReferenceId(excelData.get(i).getBColumn());
						if (StringUtils.isEmpty(excelData.get(i).getOColumn())) { // if กิจกรรมความเสียหาย เป็นค่าว่าง
																					// then
																					// Loss Profile
							incidentRow.setIncidentTitle(excelData.get(i).getAxColumn());
						} else {
							incidentRow.setIncidentTitle(excelData.get(i).getOColumn());
						}
						incidentRow.setDiscoveredBy("อื่น ๆ (Other)");
						incidentRow.setDescription(
								excelData.get(i).getAbColumn().replace("\n", "<br/>").replace("\r", "<br/>"));
						incidentRow.setDiscoveryDetail("");
						incidentRow.setMoLevel2(excelData.get(i).getBcColumn());
						incidentRow.setMoLevel3(excelData.get(i).getBdColumn());
						incidentRow.setMoLevel4(excelData.get(i).getBeColumn());
						incidentRow.setMoLevel5(excelData.get(i).getBfColumn());
						incidentRow.setMoLevel6(excelData.get(i).getBgColumn());
						incidentRow.setFirstDateOfEvent(excelData.get(i).getTColumn());
						incidentRow.setRiskEventType(excelData.get(i).getBhColumn());
						incidentRow.setAmount("0");
						incidentRow.setDateOfDiscovery(excelData.get(i).getVColumn());
						incidentRow.setLastDateOfEvent(excelData.get(i).getUColumn());
						incidentRow.setClosedDate(excelData.get(i).getAjColumn());
						incidentRow.setSourceOfIncident("OpLoss Incident");
						incidentRow.setContactDetail("");
						incidentRow.setIncidentDetailsComments("");
						incidentRow.setCustomerComplaint("N");
						incidentRow.setHasLegalImpact("N");
						incidentRow.setComplaintId("");
						incidentRow.setLegalCaseId("");
						incidentRow.setLegalCaseStatus("");
						incidentRow.setRiskCategories(excelData.get(i).getBiColumn());
						if (StringUtils.isEmpty(excelData.get(i).getAnColumn())) {
							incidentRow.setCreditRelatedOrMarketRelated("None");
						} else {
							incidentRow.setCreditRelatedOrMarketRelated(excelData.get(i).getAnColumn());
						}
						incidentRow.setProcess(excelData.get(i).getBjColumn());
						incidentRow.setProduct(excelData.get(i).getBkColumn());
						incidentRow.setCause1(excelData.get(i).getBlColumn());
						incidentRow.setCauseDescription(
								excelData.get(i).getAcColumn().replace("\n", "<br/>").replace("\r", "<br/>"));
						incidentRow.setControlDetails(
								excelData.get(i).getArColumn().replace("\n", "<br/>").replace("\r", "<br/>"));
						if (StringUtils.isEmpty(incidentRow.getControlDetails())) {
							incidentRow.setCorrectiveAction("-");
						} else {
							incidentRow.setCorrectiveAction(incidentRow.getControlDetails());
						}
						incidentModels.add(incidentRow);
						relatedOpLossIncident.append(excelData.get(i).getBColumn()).append(";");

						FinancialModel financialRow = new FinancialModel();
						financialRow.setOpLossReferenceId(excelData.get(i).getBColumn());
						financialRow.setFinancialStatus(excelData.get(i).getBmColumn());
						BigDecimal GROSS_LOSS_AMT = !StringUtils.isEmpty(excelData.get(i).getAeColumn())
							    ? new BigDecimal(excelData.get(i).getAeColumn())  // Use BigDecimal for precision
							    : BigDecimal.ZERO;
						if (GROSS_LOSS_AMT.compareTo(BigDecimal.ZERO) < 0) { // If less than 0
						    financialRow.setEffectNature("Gain");
						    financialRow.setEffectLossAmount("");
						    financialRow.setEffectGainAmount(GROSS_LOSS_AMT.toPlainString()); // Avoid scientific notation
						} else {
						    financialRow.setEffectNature("Loss");
						    financialRow.setEffectLossAmount(GROSS_LOSS_AMT.toPlainString()); // Avoid scientific notation
						    financialRow.setEffectGainAmount("");
						}
						financialRow.setBookingDate(excelData.get(i).getAfColumn());
						if (!StringUtils.isEmpty(excelData.get(i).getAfColumn())
								&& excelData.get(i).getAaColumn().equals("ยังไม่บันทึกบัญชี หรือไม่สามารถระบุได้")) {
							financialRow.setLedgerCode("ยังไม่บันทึกบัญชี");
						} else {
							financialRow.setLedgerCode(excelData.get(i).getAaColumn());
						}
						financialRow.setBookingUnitCode(excelData.get(i).getCColumn());
						financialRow.setBookingUnitName(excelData.get(i).getKColumn());
						financialRow.setBusinessLineCodeBusinessLine(excelData.get(i).getLColumn());
						financialModels.add(financialRow);
					} else {
						IncidentModel incidentRow = new IncidentModel();
						incidentRow.setReferenceId(excelData.get(i).getBColumn());
						if (StringUtils.isEmpty(excelData.get(i).getOColumn())) { // if กิจกรรมความเสียหาย เป็นค่าว่าง
																					// then Loss Profile
							incidentRow.setIncidentTitle(excelData.get(i).getAxColumn());
						} else {
							incidentRow.setIncidentTitle(excelData.get(i).getOColumn());
						}
						incidentRow.setDiscoveredBy("อื่น ๆ (Other)");
						incidentRow.setDescription(
								excelData.get(i).getAbColumn().replace("\n", "<br/>").replace("\r", "<br/>"));
						incidentRow.setDiscoveryDetail("");
						incidentRow.setMoLevel2(excelData.get(i).getBcColumn());
						incidentRow.setMoLevel3(excelData.get(i).getBdColumn());
						incidentRow.setMoLevel4(excelData.get(i).getBeColumn());
						incidentRow.setMoLevel5(excelData.get(i).getBfColumn());
						incidentRow.setMoLevel6(excelData.get(i).getBgColumn());
						incidentRow.setFirstDateOfEvent(excelData.get(i).getTColumn());
						incidentRow.setRiskEventType(excelData.get(i).getBhColumn());
						incidentRow.setAmount("0");
						incidentRow.setDateOfDiscovery(excelData.get(i).getVColumn());
						incidentRow.setLastDateOfEvent(excelData.get(i).getUColumn());
						incidentRow.setClosedDate(excelData.get(i).getAjColumn());
						incidentRow.setSourceOfIncident("OpLoss Incident");
						incidentRow.setContactDetail("");
						incidentRow.setIncidentDetailsComments("");
						incidentRow.setCustomerComplaint("N");
						incidentRow.setHasLegalImpact("N");
						incidentRow.setComplaintId("");
						incidentRow.setLegalCaseId("");
						incidentRow.setLegalCaseStatus("");
						incidentRow.setRiskCategories(excelData.get(i).getBiColumn());
						if (StringUtils.isEmpty(excelData.get(i).getAnColumn())) {
							incidentRow.setCreditRelatedOrMarketRelated("None");
						} else {
							incidentRow.setCreditRelatedOrMarketRelated(excelData.get(i).getAnColumn());
						}
						incidentRow.setProcess(excelData.get(i).getBjColumn());
						incidentRow.setProduct(excelData.get(i).getBkColumn());
						incidentRow.setCause1(excelData.get(i).getBlColumn());
						incidentRow.setCauseDescription(
								excelData.get(i).getAcColumn().replace("\n", "<br/>").replace("\r", "<br/>"));
						incidentRow.setControlDetails(
								excelData.get(i).getArColumn().replace("\n", "<br/>").replace("\r", "<br/>"));
						if (StringUtils.isEmpty(incidentRow.getControlDetails())) {
							incidentRow.setCorrectiveAction("-");
						} else {
							incidentRow.setCorrectiveAction(incidentRow.getControlDetails());
						}
						// Remove the last semicolon if it exists
						if (relatedOpLossIncident.length() > 0
								&& relatedOpLossIncident.charAt(relatedOpLossIncident.length() - 1) == ';') {
							relatedOpLossIncident.setLength(relatedOpLossIncident.length() - 1); // Remove last
																									// semicolon
						}

						// Set the relatedOpLossIncident
						incidentRow.setRelatedOpLossIncident(relatedOpLossIncident.toString());
						incidentModels.add(incidentRow);
						relatedOpLossIncident = new StringBuilder();

						FinancialModel financialRow = new FinancialModel();
						financialRow.setOpLossReferenceId(excelData.get(i).getBColumn());
						financialRow.setFinancialStatus(excelData.get(i).getBmColumn());
						BigDecimal GROSS_LOSS_AMT = !StringUtils.isEmpty(excelData.get(i).getAeColumn())
							    ? new BigDecimal(excelData.get(i).getAeColumn())  // Use BigDecimal for precision
							    : BigDecimal.ZERO;
						if (GROSS_LOSS_AMT.compareTo(BigDecimal.ZERO) < 0) { // If less than 0
						    financialRow.setEffectNature("Gain");
						    financialRow.setEffectLossAmount("");
						    financialRow.setEffectGainAmount(GROSS_LOSS_AMT.toPlainString()); // Avoid scientific notation
						} else {
						    financialRow.setEffectNature("Loss");
						    financialRow.setEffectLossAmount(GROSS_LOSS_AMT.toPlainString()); // Avoid scientific notation
						    financialRow.setEffectGainAmount("");
						}
						financialRow.setBookingDate(excelData.get(i).getAfColumn());
						if (!StringUtils.isEmpty(excelData.get(i).getAfColumn())
								&& excelData.get(i).getAaColumn().equals("ยังไม่บันทึกบัญชี หรือไม่สามารถระบุได้")) {
							financialRow.setLedgerCode("ยังไม่บันทึกบัญชี");
						} else {
							financialRow.setLedgerCode(excelData.get(i).getAaColumn());
						}
						financialRow.setBookingUnitCode(excelData.get(i).getCColumn());
						financialRow.setBookingUnitName(excelData.get(i).getKColumn());
						financialRow.setBusinessLineCodeBusinessLine(excelData.get(i).getLColumn());
						financialModels.add(financialRow);
					}

				} else if (!changeSequence && excelData.get(i).getBColumn().equals(excelData.get(i + 1).getBColumn())) {
					// create Financial

					FinancialModel financialRow = new FinancialModel();
					financialRow.setOpLossReferenceId(excelData.get(i).getBColumn());
					financialRow.setFinancialStatus(excelData.get(i).getBmColumn());
					BigDecimal GROSS_LOSS_AMT = !StringUtils.isEmpty(excelData.get(i).getAeColumn())
						    ? new BigDecimal(excelData.get(i).getAeColumn())  // Use BigDecimal for precision
						    : BigDecimal.ZERO;
					if (GROSS_LOSS_AMT.compareTo(BigDecimal.ZERO) < 0) { // If less than 0
					    financialRow.setEffectNature("Gain");
					    financialRow.setEffectLossAmount("");
					    financialRow.setEffectGainAmount(GROSS_LOSS_AMT.toPlainString()); // Avoid scientific notation
					} else {
					    financialRow.setEffectNature("Loss");
					    financialRow.setEffectLossAmount(GROSS_LOSS_AMT.toPlainString()); // Avoid scientific notation
					    financialRow.setEffectGainAmount("");
					}
					financialRow.setBookingDate(excelData.get(i).getAfColumn());
					if (!StringUtils.isEmpty(excelData.get(i).getAfColumn())
							&& excelData.get(i).getAaColumn().equals("ยังไม่บันทึกบัญชี หรือไม่สามารถระบุได้")) {
						financialRow.setLedgerCode("ยังไม่บันทึกบัญชี");
					} else {
						financialRow.setLedgerCode(excelData.get(i).getAaColumn());
					}
					financialRow.setBookingUnitCode(excelData.get(i).getCColumn());
					financialRow.setBookingUnitName(excelData.get(i).getKColumn());
					financialRow.setBusinessLineCodeBusinessLine(excelData.get(i).getLColumn());
					financialModels.add(financialRow);
				} else {
					FinancialModel financialRow = new FinancialModel();
					financialRow.setOpLossReferenceId(excelData.get(i).getBColumn());
					financialRow.setFinancialStatus(excelData.get(i).getBmColumn());
					BigDecimal GROSS_LOSS_AMT = !StringUtils.isEmpty(excelData.get(i).getAeColumn())
						    ? new BigDecimal(excelData.get(i).getAeColumn())  // Use BigDecimal for precision
						    : BigDecimal.ZERO;
					if (GROSS_LOSS_AMT.compareTo(BigDecimal.ZERO) < 0) { // If less than 0
					    financialRow.setEffectNature("Gain");
					    financialRow.setEffectLossAmount("");
					    financialRow.setEffectGainAmount(GROSS_LOSS_AMT.toPlainString()); // Avoid scientific notation
					} else {
					    financialRow.setEffectNature("Loss");
					    financialRow.setEffectLossAmount(GROSS_LOSS_AMT.toPlainString()); // Avoid scientific notation
					    financialRow.setEffectGainAmount("");
					}
					financialRow.setBookingDate(excelData.get(i).getAfColumn());
					if (!StringUtils.isEmpty(excelData.get(i).getAfColumn())
							&& excelData.get(i).getAaColumn().equals("ยังไม่บันทึกบัญชี หรือไม่สามารถระบุได้")) {
						financialRow.setLedgerCode("ยังไม่บันทึกบัญชี");
					} else {
						financialRow.setLedgerCode(excelData.get(i).getAaColumn());
					}
					financialRow.setBookingUnitCode(excelData.get(i).getCColumn());
					financialRow.setBookingUnitName(excelData.get(i).getKColumn());
					financialRow.setBusinessLineCodeBusinessLine(excelData.get(i).getLColumn());
					financialModels.add(financialRow);
				}
			} else if (!changeSequence) { // last record
				if (!excelData.get(i).getAsColumn().equals("1.0")) {
					IncidentModel incidentRow = new IncidentModel();
					incidentRow.setReferenceId(excelData.get(i).getBColumn());
					if (StringUtils.isEmpty(excelData.get(i).getOColumn())) { // if กิจกรรมความเสียหาย เป็นค่าว่าง then
																				// Loss Profile
						incidentRow.setIncidentTitle(excelData.get(i).getAxColumn());
					} else {
						incidentRow.setIncidentTitle(excelData.get(i).getOColumn());
					}
					incidentRow.setDiscoveredBy("อื่น ๆ (Other)");
					incidentRow.setDescription(
							excelData.get(i).getAbColumn().replace("\n", "<br/>").replace("\r", "<br/>"));
					incidentRow.setDiscoveryDetail("");
					incidentRow.setMoLevel2(excelData.get(i).getBcColumn());
					incidentRow.setMoLevel3(excelData.get(i).getBdColumn());
					incidentRow.setMoLevel4(excelData.get(i).getBeColumn());
					incidentRow.setMoLevel5(excelData.get(i).getBfColumn());
					incidentRow.setMoLevel6(excelData.get(i).getBgColumn());
					incidentRow.setFirstDateOfEvent(excelData.get(i).getTColumn());
					incidentRow.setRiskEventType(excelData.get(i).getBhColumn());
					incidentRow.setAmount("0");
					incidentRow.setDateOfDiscovery(excelData.get(i).getVColumn());
					incidentRow.setLastDateOfEvent(excelData.get(i).getUColumn());
					incidentRow.setClosedDate(excelData.get(i).getAjColumn());
					incidentRow.setSourceOfIncident("OpLoss Incident");
					incidentRow.setContactDetail("");
					incidentRow.setIncidentDetailsComments("");
					incidentRow.setCustomerComplaint("N");
					incidentRow.setHasLegalImpact("N");
					incidentRow.setComplaintId("");
					incidentRow.setLegalCaseId("");
					incidentRow.setLegalCaseStatus("");
					incidentRow.setRiskCategories(excelData.get(i).getBiColumn());
					if (StringUtils.isEmpty(excelData.get(i).getAnColumn())) {
						incidentRow.setCreditRelatedOrMarketRelated("None");
					} else {
						incidentRow.setCreditRelatedOrMarketRelated(excelData.get(i).getAnColumn());
					}
					incidentRow.setProcess(excelData.get(i).getBjColumn());
					incidentRow.setProduct(excelData.get(i).getBkColumn());
					incidentRow.setCause1(excelData.get(i).getBlColumn());
					incidentRow.setCauseDescription(
							excelData.get(i).getAcColumn().replace("\n", "<br/>").replace("\r", "<br/>"));
					incidentRow.setControlDetails(
							excelData.get(i).getArColumn().replace("\n", "<br/>").replace("\r", "<br/>"));
					if (StringUtils.isEmpty(incidentRow.getControlDetails())) {
						incidentRow.setCorrectiveAction("-");
					} else {
						incidentRow.setCorrectiveAction(incidentRow.getControlDetails());
					}
					incidentModels.add(incidentRow);

					FinancialModel financialRow = new FinancialModel();
					financialRow.setOpLossReferenceId(excelData.get(i).getBColumn());
					financialRow.setFinancialStatus(excelData.get(i).getBmColumn());
					BigDecimal GROSS_LOSS_AMT = !StringUtils.isEmpty(excelData.get(i).getAeColumn())
						    ? new BigDecimal(excelData.get(i).getAeColumn())  // Use BigDecimal for precision
						    : BigDecimal.ZERO;
					if (GROSS_LOSS_AMT.compareTo(BigDecimal.ZERO) < 0) { // If less than 0
					    financialRow.setEffectNature("Gain");
					    financialRow.setEffectLossAmount("");
					    financialRow.setEffectGainAmount(GROSS_LOSS_AMT.toPlainString()); // Avoid scientific notation
					} else {
					    financialRow.setEffectNature("Loss");
					    financialRow.setEffectLossAmount(GROSS_LOSS_AMT.toPlainString()); // Avoid scientific notation
					    financialRow.setEffectGainAmount("");
					}
					financialRow.setBookingDate(excelData.get(i).getAfColumn());
					if (!StringUtils.isEmpty(excelData.get(i).getAfColumn())
							&& excelData.get(i).getAaColumn().equals("ยังไม่บันทึกบัญชี หรือไม่สามารถระบุได้")) {
						financialRow.setLedgerCode("ยังไม่บันทึกบัญชี");
					} else {
						financialRow.setLedgerCode(excelData.get(i).getAaColumn());
					}
					financialRow.setBookingUnitCode(excelData.get(i).getCColumn());
					financialRow.setBookingUnitName(excelData.get(i).getKColumn());
					financialRow.setBusinessLineCodeBusinessLine(excelData.get(i).getLColumn());
					financialModels.add(financialRow);
				} else {
					IncidentModel incidentRow = new IncidentModel();
					incidentRow.setReferenceId(excelData.get(i).getBColumn());
					if (StringUtils.isEmpty(excelData.get(i).getOColumn())) { // if กิจกรรมความเสียหาย เป็นค่าว่าง then
																				// Loss Profile
						incidentRow.setIncidentTitle(excelData.get(i).getAxColumn());
					} else {
						incidentRow.setIncidentTitle(excelData.get(i).getOColumn());
					}
					incidentRow.setDiscoveredBy("อื่น ๆ (Other)");
					incidentRow.setDescription(
							excelData.get(i).getAbColumn().replace("\n", "<br/>").replace("\r", "<br/>"));
					incidentRow.setDiscoveryDetail("");
					incidentRow.setMoLevel2(excelData.get(i).getBcColumn());
					incidentRow.setMoLevel3(excelData.get(i).getBdColumn());
					incidentRow.setMoLevel4(excelData.get(i).getBeColumn());
					incidentRow.setMoLevel5(excelData.get(i).getBfColumn());
					incidentRow.setMoLevel6(excelData.get(i).getBgColumn());
					incidentRow.setFirstDateOfEvent(excelData.get(i).getTColumn());
					incidentRow.setRiskEventType(excelData.get(i).getBhColumn());
					incidentRow.setAmount("0");
					incidentRow.setDateOfDiscovery(excelData.get(i).getVColumn());
					incidentRow.setLastDateOfEvent(excelData.get(i).getUColumn());
					incidentRow.setClosedDate(excelData.get(i).getAjColumn());
					incidentRow.setSourceOfIncident("OpLoss Incident");
					incidentRow.setContactDetail("");
					incidentRow.setIncidentDetailsComments("");
					incidentRow.setCustomerComplaint("N");
					incidentRow.setHasLegalImpact("N");
					incidentRow.setComplaintId("");
					incidentRow.setLegalCaseId("");
					incidentRow.setLegalCaseStatus("");
					incidentRow.setRiskCategories(excelData.get(i).getBiColumn());
					if (StringUtils.isEmpty(excelData.get(i).getAnColumn())) {
						incidentRow.setCreditRelatedOrMarketRelated("None");
					} else {
						incidentRow.setCreditRelatedOrMarketRelated(excelData.get(i).getAnColumn());
					}
					incidentRow.setProcess(excelData.get(i).getBjColumn());
					incidentRow.setProduct(excelData.get(i).getBkColumn());
					incidentRow.setCause1(excelData.get(i).getBlColumn());
					incidentRow.setCauseDescription(
							excelData.get(i).getAcColumn().replace("\n", "<br/>").replace("\r", "<br/>"));
					incidentRow.setControlDetails(
							excelData.get(i).getArColumn().replace("\n", "<br/>").replace("\r", "<br/>"));
					if (StringUtils.isEmpty(incidentRow.getControlDetails())) {
						incidentRow.setCorrectiveAction("-");
					} else {
						incidentRow.setCorrectiveAction(incidentRow.getControlDetails());
					}
					// Remove the last semicolon if it exists
					if (relatedOpLossIncident.length() > 0
							&& relatedOpLossIncident.charAt(relatedOpLossIncident.length() - 1) == ';') {
						relatedOpLossIncident.setLength(relatedOpLossIncident.length() - 1); // Remove last semicolon
					}

					// Set the relatedOpLossIncident
					incidentRow.setRelatedOpLossIncident(relatedOpLossIncident.toString());
					incidentModels.add(incidentRow);
					relatedOpLossIncident = new StringBuilder();

					FinancialModel financialRow = new FinancialModel();
					financialRow.setOpLossReferenceId(excelData.get(i).getBColumn());
					financialRow.setFinancialStatus(excelData.get(i).getBmColumn());
					BigDecimal GROSS_LOSS_AMT = !StringUtils.isEmpty(excelData.get(i).getAeColumn())
						    ? new BigDecimal(excelData.get(i).getAeColumn())  // Use BigDecimal for precision
						    : BigDecimal.ZERO;
					if (GROSS_LOSS_AMT.compareTo(BigDecimal.ZERO) < 0) { // If less than 0
					    financialRow.setEffectNature("Gain");
					    financialRow.setEffectLossAmount("");
					    financialRow.setEffectGainAmount(GROSS_LOSS_AMT.toPlainString()); // Avoid scientific notation
					} else {
					    financialRow.setEffectNature("Loss");
					    financialRow.setEffectLossAmount(GROSS_LOSS_AMT.toPlainString()); // Avoid scientific notation
					    financialRow.setEffectGainAmount("");
					}
					financialRow.setBookingDate(excelData.get(i).getAfColumn());
					if (!StringUtils.isEmpty(excelData.get(i).getAfColumn())
							&& excelData.get(i).getAaColumn().equals("ยังไม่บันทึกบัญชี หรือไม่สามารถระบุได้")) {
						financialRow.setLedgerCode("ยังไม่บันทึกบัญชี");
					} else {
						financialRow.setLedgerCode(excelData.get(i).getAaColumn());
					}
					financialRow.setBookingUnitCode(excelData.get(i).getCColumn());
					financialRow.setBookingUnitName(excelData.get(i).getKColumn());
					financialRow.setBusinessLineCodeBusinessLine(excelData.get(i).getLColumn());
					financialModels.add(financialRow);
				}
			}

		}

//		List<ExcelModel> filteredData = excelData.stream()
//		        .filter(model -> model.getAhColumn() != null && Integer.parseInt(model.getAhColumn()) != 0)
//		        .collect(Collectors.toList());

//		RecoveryModel
//		String tmpRef="";
//		List<String> listRef = new ArrayList<>();
//		for(int i =0;i<excelData.size();i++) {
//			if(excelData.get(i).getBColumn().equals(tmpRef)) {
//				listRef.add(tmpRef);
//			}
//		}
//		for(int i =0;i<excelData.size();i++) {
//			if(listRef.contains(excelData.get(i).getBColumn())&&!excelData.get(i).getAhColumn().equals("1.0")) {
//				//delete this excelData.get(i)
//			}
//		}

		List<String> duplicateRefs = new ArrayList<>();
		String tmpRef = "";

		// Step 1: Find duplicates in BColumn
		for (int i = 0; i < excelData.size(); i++) {
			String currentBColumn = excelData.get(i).getBColumn();
			if (tmpRef.equals(currentBColumn)) {
				duplicateRefs.add(currentBColumn); // Add to duplicates list
			} else {
				tmpRef = currentBColumn; // Update tmpRef for comparison
			}
		}

		// Step 2: Remove unwanted records where BColumn is in duplicates and AsColumn
		// != "1.0"
		excelData.removeIf(data -> (duplicateRefs.contains(data.getBColumn()) && !data.getAsColumn().equals("1.0")));
		excelData.removeIf(data -> {
			String ahColumn = data.getAhColumn();
			// Check if the string is null or empty
			if (StringUtils.isEmpty(ahColumn)) {
				return true; // Remove if it's empty
			}

			try {
				// Parse the value as a double
				double value = Double.parseDouble(ahColumn);
				// Check if the parsed value is 0
				return value == 0;
			} catch (NumberFormatException e) {
				// If the value is not numeric, remove the element
				return true;
			}
		});

		for (int i = 0; i < excelData.size(); i++) {

			BigDecimal RECOVERY_AMT = !StringUtils.isEmpty(excelData.get(i).getAhColumn()) 
                    ? new BigDecimal(excelData.get(i).getAhColumn()).setScale(2, RoundingMode.HALF_UP)
                    : BigDecimal.ZERO;
			    
			    BigDecimal INSURANCE_RECOVERY_AMT = !StringUtils.isEmpty(excelData.get(i).getAlColumn()) 
			        ? new BigDecimal(excelData.get(i).getAlColumn()).setScale(2, RoundingMode.HALF_UP)
			        : BigDecimal.ZERO;
			String result;
			if (RECOVERY_AMT.compareTo(BigDecimal.ZERO) < 0) {
				result = "ไม่สามารถระบุได้ (Cannot be identified)";
				RecoveryModel recoveryRecord = new RecoveryModel();
				recoveryRecord.setRecoveryType(result);
				recoveryRecord.setRecoveryAmount(RECOVERY_AMT.toPlainString());
				recoveryRecord.setRecoveryDescription(
						excelData.get(i).getAbColumn().replace("\n", "<br/>").replace("\r", "<br/>"));
				recoveryRecord.setBookingDate(excelData.get(i).getAfColumn());
				if (!StringUtils.isEmpty(excelData.get(i).getAfColumn())
						&& excelData.get(i).getAaColumn().equals("ยังไม่บันทึกบัญชี หรือไม่สามารถระบุได้")) {
					recoveryRecord.setLedgerCode("ยังไม่บันทึกบัญชี");
				} else {
					recoveryRecord.setLedgerCode(excelData.get(i).getAaColumn());
				}
				recoveryRecord.setBookingUnitCode(excelData.get(i).getCColumn());
				recoveryRecord.setBookingUnitName(excelData.get(i).getKColumn());
				recoveryRecord.setOpLossReferenceId(excelData.get(i).getBColumn());
				recoveryModels.add(recoveryRecord);
			} else if (RECOVERY_AMT.compareTo(BigDecimal.ZERO) > 0 && INSURANCE_RECOVERY_AMT.compareTo(BigDecimal.ZERO) == 0) {
				result = "ไม่สามารถระบุได้ (Cannot be identified)";
				RecoveryModel recoveryRecord = new RecoveryModel();
				recoveryRecord.setRecoveryType(result);
				recoveryRecord.setRecoveryAmount(RECOVERY_AMT.toPlainString());
				recoveryRecord.setRecoveryDescription(
						excelData.get(i).getAbColumn().replace("\n", "<br/>").replace("\r", "<br/>"));
				recoveryRecord.setBookingDate(excelData.get(i).getAfColumn());
				if (!StringUtils.isEmpty(excelData.get(i).getAfColumn())
						&& excelData.get(i).getAaColumn().equals("ยังไม่บันทึกบัญชี หรือไม่สามารถระบุได้")) {
					recoveryRecord.setLedgerCode("ยังไม่บันทึกบัญชี");
				} else {
					recoveryRecord.setLedgerCode(excelData.get(i).getAaColumn());
				}
				recoveryRecord.setBookingUnitCode(excelData.get(i).getCColumn());
				recoveryRecord.setBookingUnitName(excelData.get(i).getKColumn());
				recoveryRecord.setOpLossReferenceId(excelData.get(i).getBColumn());
				recoveryModels.add(recoveryRecord);
			} else if (RECOVERY_AMT.compareTo(BigDecimal.ZERO) > 0 && INSURANCE_RECOVERY_AMT.compareTo(BigDecimal.ZERO) > 0 &&
		               INSURANCE_RECOVERY_AMT.compareTo(RECOVERY_AMT) == 0) {
				result = "ค่าสินไหมทดแทนจากประกันภัย (Claims from insurance)";
				RecoveryModel recoveryRecord = new RecoveryModel();
				recoveryRecord.setRecoveryType(result);
				recoveryRecord.setRecoveryAmount(RECOVERY_AMT.toPlainString());
				recoveryRecord.setRecoveryDescription(
						excelData.get(i).getAbColumn().replace("\n", "<br/>").replace("\r", "<br/>"));
				recoveryRecord.setBookingDate(excelData.get(i).getAfColumn());
				if (!StringUtils.isEmpty(excelData.get(i).getAfColumn())
						&& excelData.get(i).getAaColumn().equals("ยังไม่บันทึกบัญชี หรือไม่สามารถระบุได้")) {
					recoveryRecord.setLedgerCode("ยังไม่บันทึกบัญชี");
				} else {
					recoveryRecord.setLedgerCode(excelData.get(i).getAaColumn());
				}
				recoveryRecord.setBookingUnitCode(excelData.get(i).getCColumn());
				recoveryRecord.setBookingUnitName(excelData.get(i).getKColumn());
				recoveryRecord.setOpLossReferenceId(excelData.get(i).getBColumn());
				recoveryModels.add(recoveryRecord);
			} else if (RECOVERY_AMT.compareTo(BigDecimal.ZERO) > 0 && INSURANCE_RECOVERY_AMT.compareTo(BigDecimal.ZERO) > 0 &&
		               INSURANCE_RECOVERY_AMT.compareTo(RECOVERY_AMT) < 0) {
				RecoveryModel recoveryRecord = new RecoveryModel();
				recoveryRecord.setRecoveryType("ไม่สามารถระบุได้ (Cannot be identified)");
				recoveryRecord.setRecoveryAmount(RECOVERY_AMT.subtract(INSURANCE_RECOVERY_AMT).toPlainString());
				recoveryRecord.setRecoveryDescription(
						excelData.get(i).getAbColumn().replace("\n", "<br/>").replace("\r", "<br/>"));
				recoveryRecord.setBookingDate(excelData.get(i).getAfColumn());
				if (!StringUtils.isEmpty(excelData.get(i).getAfColumn())
						&& excelData.get(i).getAaColumn().equals("ยังไม่บันทึกบัญชี หรือไม่สามารถระบุได้")) {
					recoveryRecord.setLedgerCode("ยังไม่บันทึกบัญชี");
				} else {
					recoveryRecord.setLedgerCode(excelData.get(i).getAaColumn());
				}
				recoveryRecord.setBookingUnitCode(excelData.get(i).getCColumn());
				recoveryRecord.setBookingUnitName(excelData.get(i).getKColumn());
				recoveryRecord.setOpLossReferenceId(excelData.get(i).getBColumn());
				recoveryModels.add(recoveryRecord);

				recoveryRecord = new RecoveryModel();
				recoveryRecord.setRecoveryType("ค่าสินไหมทดแทนจากประกันภัย (Claims from insurance)");
				recoveryRecord.setRecoveryAmount(INSURANCE_RECOVERY_AMT.toPlainString());
				recoveryRecord.setRecoveryDescription(
						excelData.get(i).getAbColumn().replace("\n", "<br/>").replace("\r", "<br/>"));
				recoveryRecord.setBookingDate(excelData.get(i).getAfColumn());
				if (!StringUtils.isEmpty(excelData.get(i).getAfColumn())
						&& excelData.get(i).getAaColumn().equals("ยังไม่บันทึกบัญชี หรือไม่สามารถระบุได้")) {
					recoveryRecord.setLedgerCode("ยังไม่บันทึกบัญชี");
				} else {
					recoveryRecord.setLedgerCode(excelData.get(i).getAaColumn());
				}
				recoveryRecord.setBookingUnitCode(excelData.get(i).getCColumn());
				recoveryRecord.setBookingUnitName(excelData.get(i).getKColumn());
				recoveryRecord.setOpLossReferenceId(excelData.get(i).getBColumn());
				recoveryModels.add(recoveryRecord);

			}

//			System.out.println(i+" "+excelData.get(i).getBColumn());
		}

//		Map<String, List<ExcelModel>> groupedByBColumn = excelData.stream()
//                .filter(model -> model.getBColumn() != null) // Ensure BColumn is not null
//                .collect(Collectors.groupingBy(ExcelModel::getBColumn, Collectors.toList()));
//
//        // Step 2: Filter out groups where any record has AsColumn != "1.0"
//        List<ExcelModel> filteredData = groupedByBColumn.entrySet().stream()
//                .filter(entry -> entry.getValue().stream() // Keep groups where all records have AsColumn = "1.0"
//                        .allMatch(model -> "1.0".equals(model.getAsColumn())))
//                .flatMap(entry -> entry.getValue().stream()) // Flatten the remaining groups
//                .collect(Collectors.toList());
//        for(int i =0;i<filteredData.size();i++) {
//			System.out.println(filteredData.get(i).getBColumn());
//		}

		exportIncidentCSV("C:/Users/Toonch/Downloads", incidentModels);
		exportFinancialCSV("C:/Users/Toonch/Downloads", financialModels);
		exportRecoveryCSV("C:/Users/Toonch/Downloads", recoveryModels);
		System.out.println("End");

	}

	public void exportIncidentCSV(String csvFilePath, List<IncidentModel> incidentModels) {

		try (BufferedWriter writer = new BufferedWriter(new OutputStreamWriter(
				new FileOutputStream((csvFilePath) + java.io.File.separator + "Incident.csv"), StandardCharsets.UTF_8));
//				CustomCSVWriter csvWriter = new CustomCSVWriter(writer, ',', '"', '\\', System.lineSeparator())
				CSVWriter csvWriter = new CSVWriter(writer, '|', CSVWriter.NO_QUOTE_CHARACTER, '\\',
						System.lineSeparator())) {

			// Write the header
			String[] header = { "Reference ID", "Related OpLoss Incident (Reference ID)", "Incident Title",
					"Discovered By", "Description", "Discovery Detail", "MO Level2 (MO Level2 Code)",
					"MO Level3 (MO Level3 Code)", "MO Level4 (MO Level4 Code)", "MO Level5 (MO Level5 Code)",
					"MO Level6 (MO Level6 Code)", "First Date of Event", "Date of Discovery", "Risk Event Type",
					"Amount", "Last Date of Event", "Closed Date", "Source of Incident", "Contact Detail",
					"Incident Details Comments", "Customer Complaint?", "Has Legal Impact?", "Complaint ID",
					"Legal Case ID", "Legal Case Status", "Prosecution Results", "Days for SLA1", "Days for SLA2",
					"Risk Categories (Risk Code)", "Credit Related or Market Related", "Process",
					"Product", "Cause 1", "Cause 2",
					"Cause Description", "Failed Control", "Control Details", "Corrective Action",
					"Impact Type", "Brand Impact Classification", "Employee Impact Classification",
					"Environmental Impact Classification", "Legal Impact Classification",
					"Regulatory Impact Classification", "Reputational Impact Classification",
					"Stakeholder Impact Classification", "Brand Impact Description", "Employee Impact Description",
					"Environmental Impact Description", "Legal Impact Description", "Regulatory Impact Description",
					"Reputational Impact Description", "Stakeholder Impact Description", "Issues (Issue Number)",
					"Reference ID(ITSM Interface / Bulk Upload)", "Bulk upload reference ID","Historical Data Upload?" };
			csvWriter.writeNext(header);

			// Write the data
			for (IncidentModel incident : incidentModels) {
				String referenceId = incident.getReferenceId();
				String relatedOpLossIncident = incident.getRelatedOpLossIncident() != null
						? incident.getRelatedOpLossIncident()
						: "";
				String incidentTitle = incident.getIncidentTitle();
				String discoveredBy = incident.getDiscoveredBy();
				String description = incident.getDescription();
				String discoveryDetail = incident.getDiscoveryDetail();
				String MOLevel2Code = removeDigits(incident.getMoLevel2());
				String MOLevel3Code = removeDigits(incident.getMoLevel3());
				String MOLevel4Code = removeDigits(incident.getMoLevel4());
				String MOLevel5Code = removeDigits(incident.getMoLevel5());
				String MOLevel6Code = removeDigits(incident.getMoLevel6());
				String firstDateofEvent = incident.getFirstDateOfEvent();
				String dateofDiscovery = incident.getDateOfDiscovery();
				String riskEventType = incident.getRiskEventType();
				String amount = incident.getAmount();
				String lastDateofEvent = incident.getLastDateOfEvent();
				String closedDate = incident.getClosedDate();
				String sourceofIncident = incident.getSourceOfIncident();
				String contactDetail = incident.getContactDetail();
				String incidentDetailsComments = incident.getIncidentDetailsComments();
				String customerComplaint = incident.getCustomerComplaint();
				String hasLegalImpact = incident.getHasLegalImpact();
				String complaintID = incident.getComplaintId();
				String legalCaseID = incident.getLegalCaseId();
				String legalCaseStatus = incident.getLegalCaseStatus();
				String riskCategories = incident.getRiskCategories();
				String cROrMRRelated = incident.getCreditRelatedOrMarketRelated();
				String Process = incident.getProcess();
				String Product = incident.getProduct();
				String Cause1 = incident.getCause1();
				String causeDescription = incident.getCauseDescription();
				String controlDetails = incident.getControlDetails();
				String correctiveAction = incident.getCorrectiveAction();

				// Write each row to the CSV
				String[] row = { referenceId, relatedOpLossIncident, incidentTitle, discoveredBy, description,
						discoveryDetail, MOLevel2Code, MOLevel3Code, MOLevel4Code, MOLevel5Code, MOLevel6Code,
						firstDateofEvent, dateofDiscovery, riskEventType, amount, lastDateofEvent, closedDate,
						sourceofIncident, contactDetail, incidentDetailsComments, customerComplaint, hasLegalImpact,
						complaintID, legalCaseID, legalCaseStatus, "", "", "", riskCategories, cROrMRRelated, Process,
						Product, Cause1, "", causeDescription, "Preventive Control", controlDetails, correctiveAction,
						"", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", referenceId,
						"Proceed with Closure","Yes" };
				csvWriter.writeNext(row);
			}

			System.out.println("CSV file has been created successfully.");

		} catch (IOException e) {
			e.printStackTrace();
		}

	}

	public void exportFinancialCSV(String csvFilePath, List<FinancialModel> financialModels) {

		try (BufferedWriter writer = new BufferedWriter(
				new OutputStreamWriter(new FileOutputStream((csvFilePath) + java.io.File.separator + "Financial.csv"),
						StandardCharsets.UTF_8));
//				CustomCSVWriter csvWriter = new CustomCSVWriter(writer, ',', '"', '\\', System.lineSeparator())
				CSVWriter csvWriter = new CSVWriter(writer, '|', CSVWriter.NO_QUOTE_CHARACTER, '\\',
						System.lineSeparator())) {

			// Write the header
			String[] header = { "Tracking ID", "OpLoss (Reference ID)", "Financial Status", "Effect Nature",
					"Effect Loss Amount", "Effect Gain Amount", "Booking Date", "Ledger Code", "Other Ledger Code",
					"Reference", "Booking Unit Code", "Booking Unit Name", "Business Line (Code Business Line)" };
			csvWriter.writeNext(header);

			// Write the data
			for (FinancialModel financial : financialModels) {
				String opLossReferenceId = financial.getOpLossReferenceId();
				String financialStatus = financial.getFinancialStatus();
				String effectNature = financial.getEffectNature();
				String effectLossAmount = financial.getEffectLossAmount();
				String effectGainAmount = financial.getEffectGainAmount();
				String bookingDate = financial.getBookingDate();
				String ledgerCode = financial.getLedgerCode();
				String bookingUnitCode = financial.getBookingUnitCode().replace(".0", "");
				String bookingUnitName = financial.getBookingUnitName();
				String businessLineCodeBusinessLine = financial.getBusinessLineCodeBusinessLine();

				// Write each row to the CSV
				String[] row = { "", opLossReferenceId, financialStatus, effectNature, effectLossAmount,
						effectGainAmount, bookingDate, ledgerCode, "", "", bookingUnitCode, bookingUnitName,
						businessLineCodeBusinessLine };
				csvWriter.writeNext(row);
			}

			System.out.println("CSV file has been created successfully.");

		} catch (IOException e) {
			e.printStackTrace();
		}

	}

	public void exportRecoveryCSV(String csvFilePath, List<RecoveryModel> recoveryModels) {

		try (BufferedWriter writer = new BufferedWriter(
				new OutputStreamWriter(new FileOutputStream((csvFilePath) + java.io.File.separator + "Recovery.csv"),
						StandardCharsets.UTF_8));
//				CustomCSVWriter csvWriter = new CustomCSVWriter(writer, ',', '"', '\\', System.lineSeparator())
				CSVWriter csvWriter = new CSVWriter(writer, '|', CSVWriter.NO_QUOTE_CHARACTER, '\\',
						System.lineSeparator())) {

			// Write the header
			String[] header = { "Recovery Tracking ID", "Recovery Type", "Recovery Amount", "Recovery Description",
					"Booking Unit Code", "Booking Unit Name", "Ledger Code", "Other Ledger Code", "Booking Date",
					"OpLoss (Reference ID)" };
			csvWriter.writeNext(header);

			// Write the data
			for (RecoveryModel recovery : recoveryModels) {
				String recoveryType = recovery.getRecoveryType();
				String recoveryAmount = recovery.getRecoveryAmount();
				String recoveryDescription = recovery.getRecoveryDescription();
				String bookingUnitCode = recovery.getBookingUnitCode().replace(".0", "");
				String bookingUnitName = recovery.getBookingUnitName();
				String ledgerCode = recovery.getLedgerCode();
				String bookingDate = recovery.getBookingDate();
				String opLossReferenceId = recovery.getOpLossReferenceId();

				// Write each row to the CSV
				String[] row = { "", recoveryType, recoveryAmount, recoveryDescription, bookingUnitCode,
						bookingUnitName, ledgerCode, "", bookingDate, opLossReferenceId };
				csvWriter.writeNext(row);
			}

			System.out.println("CSV file has been created successfully.");

		} catch (IOException e) {
			e.printStackTrace();
		}

	}

	public List<ExcelModel> readExcelToModel(String filePath) {
		List<ExcelModel> excelData = new ArrayList<>();

		try (FileInputStream fis = new FileInputStream(new File(filePath));
				Workbook workbook = WorkbookFactory.create(fis)) {

			Sheet sheet = workbook.getSheetAt(0); // Assuming the first sheet contains data
			Iterator<Row> rowIterator = sheet.iterator();
			rowIterator.next(); // Skip the header row

			while (rowIterator.hasNext()) {
				Row row = rowIterator.next();
				ExcelModel model = new ExcelModel();

				// Map columns to ExcelModel fields
				model.setAColumn(getCellValue(row.getCell(0))); // Column A
				model.setBColumn(getCellValue(row.getCell(1))); // Column B
				model.setCColumn(getCellValue(row.getCell(2))); // Column C
				model.setDColumn(getCellValue(row.getCell(3))); // Column D
				model.setEColumn(getCellValue(row.getCell(4))); // Column E
				model.setFColumn(getCellValue(row.getCell(5))); // Column F
				model.setGColumn(getCellValue(row.getCell(6))); // Column G
				model.setHColumn(getCellValue(row.getCell(7))); // Column H
				model.setIColumn(getCellValue(row.getCell(8))); // Column I
				model.setJColumn(getCellValue(row.getCell(9))); // Column J
				model.setKColumn(getCellValue(row.getCell(10))); // Column K
				model.setLColumn(getCellValue(row.getCell(11))); // Column L
				model.setMColumn(getCellValue(row.getCell(12))); // Column M
				model.setNColumn(getCellValue(row.getCell(13))); // Column N
				model.setOColumn(getCellValue(row.getCell(14))); // Column O
				model.setPColumn(getCellValue(row.getCell(15))); // Column P
				model.setQColumn(getCellValue(row.getCell(16))); // Column Q
				model.setRColumn(getCellValue(row.getCell(17))); // Column R
				model.setSColumn(getCellValue(row.getCell(18))); // Column S
				model.setTColumn(getCellValue(row.getCell(19))); // Column T
				model.setUColumn(getCellValue(row.getCell(20))); // Column U
				model.setVColumn(getCellValue(row.getCell(21))); // Column V
				model.setWColumn(getCellValue(row.getCell(22))); // Column W
				model.setXColumn(getCellValue(row.getCell(23))); // Column X
				model.setYColumn(getCellValue(row.getCell(24))); // Column Y
				model.setZColumn(getCellValue(row.getCell(25))); // Column Z
				model.setAaColumn(getCellValue(row.getCell(26))); // Column AA
				model.setAbColumn(getCellValue(row.getCell(27))); // Column AB
				model.setAcColumn(getCellValue(row.getCell(28))); // Column AC
				model.setAdColumn(getCellValue(row.getCell(29))); // Column AD
				model.setAeColumn(getCellValue(row.getCell(30))); // Column AE
				model.setAfColumn(getCellValue(row.getCell(31))); // Column AF
				model.setAgColumn(getCellValue(row.getCell(32))); // Column AG
				model.setAhColumn(getCellValue(row.getCell(33))); // Column AH
				model.setAiColumn(getCellValue(row.getCell(34))); // Column AI
				model.setAjColumn(getCellValue(row.getCell(35))); // Column AJ
				model.setAkColumn(getCellValue(row.getCell(36))); // Column AK
				model.setAlColumn(getCellValue(row.getCell(37))); // Column AL
				model.setAmColumn(getCellValue(row.getCell(38))); // Column AM
				model.setAnColumn(getCellValue(row.getCell(39))); // Column AN
				model.setAoColumn(getCellValue(row.getCell(40))); // Column AO
				model.setApColumn(getCellValue(row.getCell(41))); // Column AP
				model.setAqColumn(getCellValue(row.getCell(42))); // Column AQ
				model.setArColumn(getCellValue(row.getCell(43))); // Column AR
				model.setAsColumn(getCellValue(row.getCell(44))); // Column AS
				model.setAtColumn(getCellValue(row.getCell(45))); // Column AT
				model.setAuColumn(getCellValue(row.getCell(46))); // Column AU
				model.setAvColumn(getCellValue(row.getCell(47))); // Column AV
				model.setAwColumn(getCellValue(row.getCell(48))); // Column AW
				model.setAxColumn(getCellValue(row.getCell(49))); // Column AX
				model.setAyColumn(getCellValue(row.getCell(50))); // Column AY
				model.setAzColumn(getCellValue(row.getCell(51))); // Column AZ
				model.setBaColumn(getCellValue(row.getCell(52))); // Column BA
				model.setBbColumn(getCellValue(row.getCell(53))); // Column BB
				model.setBcColumn(getCellValue(row.getCell(54))); // Column BC
				model.setBdColumn(getCellValue(row.getCell(55))); // Column BD
				model.setBeColumn(getCellValue(row.getCell(56))); // Column BE
				model.setBfColumn(getCellValue(row.getCell(57))); // Column BF
				model.setBgColumn(getCellValue(row.getCell(58))); // Column BG
				model.setBhColumn(getCellValue(row.getCell(59))); // Column BH
				model.setBiColumn(getCellValue(row.getCell(60))); // Column BI
				model.setBjColumn(getCellValue(row.getCell(61))); // Column BJ
				model.setBkColumn(getCellValue(row.getCell(62))); // Column BK
				model.setBlColumn(getCellValue(row.getCell(63))); // Column BL
				model.setBmColumn(getCellValue(row.getCell(64))); // Column BM

				excelData.add(model);
			}

		} catch (Exception e) {
			e.printStackTrace();
		}

		return excelData;
	}

	private String getCellValue(Cell cell) {
		if (cell == null) {
			return "";
		}
		switch (cell.getCellType()) {
		case STRING:
			if (cell.getStringCellValue().matches("\\d{1,2}/\\d{1,2}/\\d{4}")) {
				try {
					SimpleDateFormat inputFormat = new SimpleDateFormat("dd/MM/yyyy", Locale.ENGLISH);
					SimpleDateFormat outputFormat = new SimpleDateFormat("dd-MM-yyyy", Locale.ENGLISH);
					return outputFormat.format(inputFormat.parse(cell.getStringCellValue()));
				} catch (Exception e) {
					System.err.println("Error parsing date: ");
					return "";
				}
			} else
				return cell.getStringCellValue();
		case NUMERIC:
			if (DateUtil.isCellDateFormatted(cell)) {
				SimpleDateFormat dateFormat = new SimpleDateFormat("dd-MM-yyyy");
//				System.out.println(dateFormat.format(cell.getDateCellValue()));
				return dateFormat.format(cell.getDateCellValue()); // Return date as string
			} else {
				BigDecimal value = BigDecimal.valueOf(cell.getNumericCellValue());
		        return value.toPlainString(); // Return full number as string
			}
		case BOOLEAN:
			return String.valueOf(cell.getBooleanCellValue()); // Return boolean as string
		case FORMULA:
			// Get the evaluated value of the formula cell
			// Check the actual type of the result of the formula
			if (cell.getCachedFormulaResultType() == CellType.NUMERIC) {
				return String.valueOf(cell.getNumericCellValue()); // Formula result as string
			} else if (cell.getCachedFormulaResultType() == CellType.STRING) {
				return cell.getStringCellValue(); // Formula result as string
			} else {
				return ""; // Default case for unhandled formula results
			}
		default:
			return "";
		}
	}

	private class CustomCSVWriter extends CSVWriter {
		private boolean isHeader = true;

		public CustomCSVWriter(Writer writer, char separator, char quotechar, char escapechar, String lineEnd) {
			super(writer, separator, quotechar, escapechar, lineEnd);
		}

		@Override
		public void writeNext(String[] nextLine, boolean applyQuotesToAll) {
			if (isHeader) {
				super.writeNext(nextLine, false); // Exclude quotes for the header
				isHeader = false;
			} else {
				super.writeNext(nextLine, true); // Enclose quotes for values excluding the header
			}
		}
	}

	private String removeDigits(Object code) {
		if (code != null) {
			// Convert the code to a string
			String codeStr = code.toString();

			// If the string contains a decimal point, remove everything from the decimal to
			// the end
			if (codeStr.contains(".")) {
				return codeStr.substring(0, codeStr.indexOf("."));
			}

			return codeStr; // Otherwise, just return the original string
		}
		return ""; // Return an empty string if code is null
	}
}

package com.volksoftech.rest; 

import java.io.File; 

import javax.annotation.PostConstruct; 
import javax.servlet.ServletContext; 
import javax.servlet.http.HttpServletRequest; 
import javax.ws.rs.Consumes; 
import javax.ws.rs.GET; 
import javax.ws.rs.POST; 
import javax.ws.rs.Path; 
import javax.ws.rs.PathParam; 
import javax.ws.rs.Produces; 
import javax.ws.rs.core.Context; 
import javax.ws.rs.core.MediaType; 
import javax.ws.rs.core.Response; 
import javax.xml.parsers.DocumentBuilder; 
import javax.xml.parsers.DocumentBuilderFactory; 

import java.io.FileReader; 
import java.io.IOException; 
import java.io.InputStreamReader; 
import java.io.BufferedReader; 

import org.apache.http.message.BasicNameValuePair; 
import org.apache.poi.ss.formula.functions.Now; 
import org.apache.poi.ss.formula.functions.Today; 
import org.json.JSONObject; 

import com.volksoftech.common.Constants; 
import com.volksoftech.common.GenerateQrCode; 
import com.volksoftech.common.LogUtils;import java.util.logging.Level; 
import com.volksoftech.common.LoginEncryptionUtils; 
import com.volksoftech.common.PropertyUtils; 
import com.volksoftech.common.SessionUtils; 
import com.volksoftech.common.Utils; 
import com.volksoftech.common.WebLOSUtils; 
import com.volksoftech.common.WebServiceUtils; 
import com.volksoftech.common.WebServiceUtils.ServiceCallStatus; 
import com.volksoftech.reporting.ReportUtils; 
import com.volksoftech.security.LoggerFilter; 
import com.weblos.bll.BllCommon; 
import com.weblos.bll.BllHouseVerificationCustomers; 
import com.weblos.bll.BllLMSIntegrationService; 
import com.weblos.bll.BllSavingsLoanDocuments; 
import com.weblos.bll.BllServicePointMaster; 
import com.weblos.model.AadhaarVaultDO; 
import com.weblos.model.HttpResultDO; 
import com.weblos.model.LoanDocumentsDO; 
import com.weblos.model.RequestParamsDO; 
import com.weblos.model.RequestStatusDO; 
import com.weblos.model.ScheduleResponse; 
import com.weblos.model.WrapperRequest; 
import com.weblos.model.WrapperResponse; 
import com.weblos.model.LoanProductEnrollmentDO; 
import com.weblos.model.RepaymentSchduleDatesDO; 

import oracle.sql.DATE; 
import java.io.ByteArrayInputStream; 
import java.io.FileInputStream; 
import java.io.FileOutputStream; 
import java.io.FileWriter; 
import java.io.OutputStream; 
import java.io.StringReader; 
import java.math.BigDecimal; 
import java.net.HttpURLConnection; 
import java.net.URL; 
import java.nio.charset.Charset; 
import java.nio.charset.StandardCharsets; 
import java.nio.file.Files; 
import java.nio.file.Paths; 
import java.sql.Date; 
import java.sql.ResultSet; 
import java.sql.ResultSetMetaData; 
import java.text.DateFormat; 
import java.text.DecimalFormat; 
import java.text.SimpleDateFormat; 
import java.time.LocalDate; 
import java.time.LocalDateTime; 
import java.time.format.DateTimeFormatter; 
import java.util.ArrayList; 
import java.util.Arrays; 
import java.util.Base64; 
import java.util.Calendar; 
import java.util.List; 
import java.util.Locale; 
import java.util.Map; 
import java.util.regex.Matcher; 
import java.util.regex.Pattern; 
import java.util.stream.Collectors; 

import com.weblos.bll.BllReportService; 
import com.weblos.bll.BllSavingsLog; 

import org.apache.poi.hssf.util.HSSFColor; 
import org.apache.poi.util.XMLHelper; 
import org.apache.poi.xssf.model.SharedStringsTable; 
import org.apache.poi.xssf.usermodel.XSSFCellStyle; 
import org.codehaus.jackson.map.ObjectMapper; 
import org.codehaus.jackson.type.TypeReference; 

import com.openhtmltopdf.pdfboxout.PdfRendererBuilder; 

import org.w3c.dom.Element; 
import org.w3c.dom.Node; 
import org.w3c.dom.NodeList; 
import org.w3c.dom.html.HTMLTableCaptionElement; 
import org.xml.sax.InputSource; 

import java.sql.CallableStatement; 
import java.sql.Connection; 
import java.sql.Statement; 

import com.volksoftech.dal.DalHelper; 
import com.weblos.model.BranchMasterDO; 
import com.weblos.model.CustomerMasterDO; 
import com.weblos.model.FundFreeCalcReq; 
import com.weblos.model.FundFreeCalcRes; 
import com.weblos.model.FundFreeCalcResp; 
import com.weblos.model.ProfileReportsDO; 
import com.weblos.dal.DalCommon; 
import com.weblos.dal.DalCustomerMaster; 
import com.weblos.dal.DalLoanProductEnrollment; 
import com.weblos.dal.DalRepaymentSchdule; 
import com.weblos.dal.DalSavingsLoanDocuments; 
import com.weblos.model.ProductMasterDO; 
import com.weblos.model.RepaymentSchduleDO; 
import com.volksoftech.common.FontUtility; 

//import com.weblos.model.ProfileReportsDO; 

@Path("/savings_loan_documents") 
public class SavingsLoanDocuments 
{ 

 BllLMSIntegrationService bllLMSIntegrationService = null; 

 public SavingsLoanDocuments() 
 { 
 bllLMSIntegrationService = new BllLMSIntegrationService(); 
 } 

 @Context 
 private HttpServletRequest request; 

 @Context 
 private ServletContext servletContext; 

 BllSavingsLog bllSavingsLog = new BllSavingsLog(); 
 BllSavingsLoanDocuments bllSavingsLoanDocuments = new BllSavingsLoanDocuments(); 
 FontUtility fontUtility = new FontUtility(); 
 DalCustomerMaster dalCustomerMaster = new DalCustomerMaster(); 
 DalLoanProductEnrollment dalLoanProductEnrollment = new DalLoanProductEnrollment(); 

 public String clearOldReportsData() 
 { 
 String result = ""; 
 String contextPath = ""; 
 try 
 { 
 // contextPath = servletContext.getRealPath(File.separator); 
 contextPath = Constants.REPORTS_PATH_URL; 
 if (contextPath == null) 
 { 
 contextPath = servletContext.getRealPath(""); 
 } 
 File directory = new File(contextPath + File.separator + "REPORTS"); 
 File[] listFiles = directory.listFiles(); 
 Calendar calendar = Calendar.getInstance(); 
 calendar.add(Calendar.DAY_OF_MONTH, -6); 
 java.util.Date minDate = calendar.getTime(); 
 Date fileDate; 
 for (File file : listFiles) 
 { 
 fileDate = new Date(file.lastModified()); 
 if (minDate.after(fileDate)) 
 { 
 file.delete(); 
 } 
 } 
 } 
 catch (Exception ex) 
 { 
 Utils.handleServerException("ReportService", "ClearOldReportsData", ex.getMessage(), ex); 
 } 
 return result; 
 } 

 public String getReportImageLogoPath() 
 { 
 String result = ""; 
 try 
 { 
 String contextPath = servletContext.getRealPath(File.separator); 
 if (contextPath == null) 
 { 
 contextPath = servletContext.getRealPath(""); 
 } 
 File reportsPath = null; 
 reportsPath = new File(contextPath + File.separator + "images"); 
 result = reportsPath.getAbsolutePath(); 
 result = result + File.separator; 
 } 
 catch (Exception ex) 
 { 
 throw ex; 
 } 
 return result; 
 } 

 public String getReportHtmlString(String file) throws IOException 
 { 
 String result = ""; 
 try 
 { 
 // BufferedReader in = new BufferedReader(new FileReader(file)); 
 BufferedReader in = new BufferedReader(new InputStreamReader(new FileInputStream(file), StandardCharsets.ISO_8859_1)); 

 String line = ""; 
 try 
 { 
 while ((line = in.readLine()) != null) 
 { 
 result += line; 
 } 
 } 
 catch (IOException e) 
 { 
 e.printStackTrace(); 
 throw e; 
 } 
 in.close(); 
 } 
 catch (Exception ex) 
 { 
 throw ex; 
 } 
 return result; 
 } 

 @POST 
 @Produces(MediaType.APPLICATION_JSON) 
 @Path("/generate_annexure_data") 
 public Response generate_AnnexureData(RequestParamsDO requestParamsDO) throws Exception 
 { 
 RequestStatusDO result = new RequestStatusDO(); 
 try 
 { 
 if (SessionUtils.validateRequest(requestParamsDO, request)) 
 { 

 AesUtil aesUtil = new AesUtil(128, 1000); 
 JSONObject result1 = new JSONObject(); 
 // String reportName = "Annexure_"+Utils.getJsonProperty("groupName", 
 // requestParamsDO.jsonFilterText) + "_" +Utils.getJsonProperty("", 
 // requestParamsDO.jsonFilterText) + ".pdf"; 
 // reportName = aesUtil.decrypt(reportName); 
 // LogUtils.logMessage(Level.INFO,"SavingsLoanDocuments annexure:= Line 1" +reportName); 
 String filePath = getReportImageLogoPath(); 
 // filePath = aesUtil.decrypt(filePath); 
 LogUtils.logMessage(Level.INFO,"SavingsLoanDocuments annexure:= Line 2" + filePath); 
 String version = Utils.getJsonProperty("version", requestParamsDO.jsonFilterText); 
 version = aesUtil.decrypt(version); 
 LogUtils.logMessage(Level.INFO,"SavingsLoanDocuments annexure:= Line 3" + version); 
 String integrationStatus = Utils.getJsonProperty("integrationStatus", requestParamsDO.jsonFilterText); 
 integrationStatus = aesUtil.decrypt(integrationStatus); 
 LogUtils.logMessage(Level.INFO,"SavingsLoanDocuments annexure:= Line 4" + integrationStatus); 
 String groupName = Utils.getJsonProperty("groupName", requestParamsDO.jsonFilterText); 
 groupName = aesUtil.decrypt(groupName); 
 LogUtils.logMessage(Level.INFO,"SavingsLoanDocuments annexure:= Line 5" + groupName); 
 String stateName = Utils.getJsonProperty("stateName", requestParamsDO.jsonFilterText); 
 stateName = aesUtil.decrypt(stateName); 
 LogUtils.logMessage(Level.INFO,"SavingsLoanDocuments annexure:= Line 6" + stateName); 
 String groupId = Utils.getJsonProperty("groupId", requestParamsDO.jsonFilterText); 
 groupId = aesUtil.decrypt(groupId); 
 String groupNumber = Utils.getJsonProperty("groupNumber", requestParamsDO.jsonFilterText); 
 groupNumber = aesUtil.decrypt(groupNumber); 
 LogUtils.logMessage(Level.INFO,"SavingsLoanDocuments annexure:= Line 7" + groupId); 
 String branchName = Utils.getJsonProperty("branchName", requestParamsDO.jsonFilterText); 
 branchName = aesUtil.decrypt(branchName); 
 LogUtils.logMessage(Level.INFO,"SavingsLoanDocuments annexure:= Line 8" + branchName); 
 String regionName = Utils.getJsonProperty("regionName", requestParamsDO.jsonFilterText); 
 regionName = aesUtil.decrypt(regionName); 
 LogUtils.logMessage(Level.INFO,"SavingsLoanDocuments annexure:= Line 9" + regionName); 
 String date = Utils.getJsonProperty("proposedDisbursementDate", requestParamsDO.jsonFilterText); 
 date = aesUtil.decrypt(date); 
 LogUtils.logMessage(Level.INFO,"SavingsLoanDocuments annexure:= Line 10" + date); 
 String customerId = Utils.getJsonProperty("customerId", requestParamsDO.jsonFilterText); 
 customerId = aesUtil.decrypt(customerId); 
 LogUtils.logMessage(Level.INFO,"SavingsLoanDocuments annexure:= Line 11" + customerId); 
 String villageName1 = Utils.getJsonProperty("villageName1", requestParamsDO.jsonFilterText); 
 villageName1 = aesUtil.decrypt(villageName1); 
 String customerIds = customerId.substring(0, customerId.length() - 1); 
 LogUtils.logMessage(Level.INFO,"SavingsLoanDocuments annexure:= Line 12" + customerIds); 
 String isPDFMergingDocument = Utils.getJsonProperty("isPDFMergingDocument", requestParamsDO.jsonFilterText); 
 String reportName = "Annexure_" + groupName + "_" + groupNumber + ".pdf"; 
 LogUtils.logMessage(Level.INFO,"SavingsLoanDocuments annexure:= Line 13" + reportName); 

 String annexureData = ReportUtils 
 .generateAnnexureData("call sp_rpt_savings_annexure_document(\"" + customerIds + "\",\"" + Utils.quotedString(date) + "\")"); 
 LogUtils.logMessage(Level.INFO,"SavingsLoanDocuments annexure:= Line 13" + annexureData); 

 String htmlString = getReportHtmlString(filePath + "Annexure.txt"); 
 LogUtils.logMessage(Level.INFO,"SavingsLoanDocuments annexure:= Line 14" + filePath); 

 htmlString = htmlString.replace("#@#Place#@#", regionName); 
 LogUtils.logMessage(Level.INFO,"SavingsLoanDocuments annexure:= Line 15" + regionName); 

 htmlString = htmlString.replace("#@#GroupIntegration#@#", integrationStatus.equals("FINISHED") ? " AD " : " BD "); 
 LogUtils.logMessage(Level.INFO,"SavingsLoanDocuments annexure:= Line 16"); 

 htmlString = htmlString.replace("#@#CensusVillage#@#", villageName1); 
 LogUtils.logMessage(Level.INFO,"SavingsLoanDocuments annexure:= Line 17"); 

 htmlString = htmlString.replace("#@#District#@#", regionName); 
 LogUtils.logMessage(Level.INFO,"SavingsLoanDocuments annexure:= Line 18" + regionName); 

 htmlString = htmlString.replace("#@#State#@#", stateName); 
 LogUtils.logMessage(Level.INFO,"SavingsLoanDocuments annexure:= Line 19" + stateName); 

 htmlString = htmlString.replace("#@#Version#@#", version); 
 LogUtils.logMessage(Level.INFO,"SavingsLoanDocuments annexure:= Line 20" + version); 

 htmlString = htmlString.replace("#@#tableData#@#", annexureData); 
 LogUtils.logMessage(Level.INFO,"SavingsLoanDocuments annexure:= Line 21" + annexureData); 

 LogUtils.logMessage(Level.INFO,"SavingsLoanDocuments annexure:= Line 22 logs insert start"); 
 bllSavingsLog.insert_savings_annexure_data(requestParamsDO); 
 LogUtils.logMessage(Level.INFO,"SavingsLoanDocuments annexure:= Line 23 logs insert end"); 

 if (isPDFMergingDocument.equals("1")) 
 { 
 clearOldReportsData(); 
 result1.put("reportId", Utils.getJsonProperty("reportId", requestParamsDO.jsonFilterText)); 
 result1.put("fileName", htmlString); 
 result.setResultObject(result1.toString()); 
 LogUtils.logMessage(Level.INFO,"SavingsLoanDocuments annexure:= Line 24"); 
 } 
 else 
 { 
 clearOldReportsData(); 
 writeHtmltoPDFFile1(reportName, htmlString, "times", "English"); 
 result1.put("reportId", Utils.getJsonProperty("reportId", requestParamsDO.jsonFilterText)); 
 result1.put("fileName", getReportFullPath(reportName, true)); 
 result.setResultObject(result1.toString()); 
 LogUtils.logMessage(Level.INFO,"SavingsLoanDocuments annexure:= Line 25"); 
 } 

 } 
 else 
 { 
 WebLOSUtils.getErrorObject(result, "1011", "Invalid Session"); 
 } 
 } 
 catch (NullPointerException npe) 
 { 
 npe = new NullPointerException("Some data of this report found null values, please contact administrator."); 
 WebLOSUtils.getErrorObject(result, "1012", npe.getMessage()); 
 } 
 catch (Exception ex) 
 { 
 WebLOSUtils.getErrorObject(result, "1012", ex.getMessage()); 
 } 
 return Response.status(200).entity(result).build(); 
 } 

 @POST 
 @Produces(MediaType.APPLICATION_JSON) 
 @Path("/generate_demondpromissory_data") 
 public Response generate_Demond_Data(RequestParamsDO requestParamsDO) throws Exception 
 { 

 RequestStatusDO result = new RequestStatusDO(); 
 try 
 { 
 if (SessionUtils.validateRequest(requestParamsDO, request)) 
 { 
 JSONObject result1 = new JSONObject(); 
 String reportName = Utils.getJsonProperty("groupId", requestParamsDO.jsonFilterText) + "_" + System.currentTimeMillis() 
 + "_mLOSDemond.pdf"; 
 String filePath = getReportImageLogoPath(); 
 String version = Utils.getJsonProperty("version", requestParamsDO.jsonFilterText); 
 String integrationStatus = Utils.getJsonProperty("integrationStatus", requestParamsDO.jsonFilterText); 
 String loanAmount = Utils.getJsonProperty("loanAmount", requestParamsDO.jsonFilterText); 
 // String stateName = 
 // Utils.getJsonProperty("customerName",requestParamsDO.jsonFilterText); 
 String groupId = Utils.getJsonProperty("groupId", requestParamsDO.jsonFilterText); 
 String branchName = Utils.getJsonProperty("branchName", requestParamsDO.jsonFilterText); 
 String regionName = Utils.getJsonProperty("regionName", requestParamsDO.jsonFilterText); 
 String date = Utils.getJsonProperty("proposedDisbursementDate", requestParamsDO.jsonFilterText); 
 String customerId = Utils.getJsonProperty("customerId", requestParamsDO.jsonFilterText); 
 String customerIds = customerId.substring(0, customerId.length() - 1); 
 // String annexureData = ReportUtils.generateAnnexureData("call 
 // sp_rpt_savings_annexure_document(\""+customerIds+"\")"); 
 String custIds[] = customerIds.split(","); 
 String[] listDPNvalues = null; 
 String htmlString = ""; 
 int rows = 0; 
 for (int i = 0; i < custIds.length; i++) 
 { 
 // String annexureData = ReportUtils.generateAnnexureData("call 
 // sp_rpt_savings_annexure_document(\""+customerIds+"\")"); 

 htmlString += getReportHtmlString(filePath + "Demondpromissory.txt"); 
 listDPNvalues = ReportUtils.documentDPNData(custIds[i]); 
 htmlString = htmlString.replace("#@#Place#@#", ""); 
 htmlString = htmlString.replace("#@#Date#@#", ""); 
 if (listDPNvalues != null) 
 { 
 htmlString = htmlString.replace("#@#CustomerName#@#", listDPNvalues[0]); 
 htmlString = htmlString.replace("#@#LoanAmount#@#", listDPNvalues[3]); 
 String loanAmountInWords = listDPNvalues[3].substring(0, listDPNvalues[3].indexOf(".")); 
 htmlString = htmlString.replace("#@#Loanamountinwords#@#", convert(Long.parseLong(loanAmountInWords))); 
 System.out.println(convert(Long.parseLong(loanAmountInWords))); 
 htmlString = htmlString.replace("#@#ROI#@#", listDPNvalues[4]); 
 } 
 htmlString = htmlString.replace("#@#DisbursementDate#@#", date); 
 htmlString = htmlString.replace("#@#Version#@#", version); 
 htmlString = htmlString.replace("#@#GroupIntegration#@#", integrationStatus.equals("FINISHED") ? " AD " : " BD "); 
 if (rows == 1) 
 { 
 htmlString += " <p style='page-break-after: always;'>&nbsp;</p> "; 
 rows = 0; 
 } 
 else 
 { 
 rows++; 
 } 
 bllSavingsLog.insert_savings_demand_promisary_customers_data(requestParamsDO, custIds[i]); 
 } 

 bllSavingsLog.insert_savings_demand_promisary_data(requestParamsDO); 

 clearOldReportsData(); 
 result1.put("reportId", Utils.getJsonProperty("reportId", requestParamsDO.jsonFilterText)); 
 result1.put("fileName", htmlString); 
 result.setResultObject(result1.toString()); 
 } 
 else 
 { 
 WebLOSUtils.getErrorObject(result, "1011", "Invalid Session"); 
 } 
 } 
 catch (Exception ex) 
 { 
 WebLOSUtils.getErrorObject(result, "1012", ex.getMessage()); 
 } 
 return Response.status(200).entity(result).build(); 
 } 

 @POST 
 @Produces(MediaType.APPLICATION_JSON) 
 @Path("/generate_sanction_letter1") 
 public Response generate_sanction_Data(RequestParamsDO requestParamsDO) throws Exception 
 { 
 RequestStatusDO result = new RequestStatusDO(); 
 try 
 { 
 if (SessionUtils.validateRequest(requestParamsDO, request)) 
 { 
 AesUtil aesUtil = new AesUtil(128, 1000); 
 JSONObject result1 = new JSONObject(); 
 String reportName = Utils.getJsonProperty("sanctionId", requestParamsDO.jsonFilterText) + "_" + System.currentTimeMillis() 
 + "_SanctionLette.html"; 
 LogUtils.logMessage(Level.INFO,"SavingsLoanDocuments sanction_letter:= line 1 " + reportName); 

 String filePath = getReportImageLogoPath(); 
 LogUtils.logMessage(Level.INFO,"SavingsLoanDocuments sanction_letter:= line 2 " + filePath); 

 String version = Utils.getJsonProperty("version", requestParamsDO.jsonFilterText); 
 version = aesUtil.decrypt(version); 
 LogUtils.logMessage(Level.INFO,"SavingsLoanDocuments sanction_letter:= line 3 " + version); 

 String integrationStatus = Utils.getJsonProperty("integrationStatus", requestParamsDO.jsonFilterText); 
 integrationStatus = aesUtil.decrypt(integrationStatus); 
 LogUtils.logMessage(Level.INFO,"SavingsLoanDocuments sanction_letter:= line 4 " + integrationStatus); 

 String groupName = Utils.getJsonProperty("groupName", requestParamsDO.jsonFilterText); 
 groupName = aesUtil.decrypt(groupName); 
 LogUtils.logMessage(Level.INFO,"SavingsLoanDocuments sanction_letter:= line 5 " + groupName); 

 String stateName = Utils.getJsonProperty("stateName", requestParamsDO.jsonFilterText); 
 stateName = aesUtil.decrypt(stateName); 
 LogUtils.logMessage(Level.INFO,"SavingsLoanDocuments sanction_letter:= line 6 " + stateName); 

 String groupId = Utils.getJsonProperty("groupId", requestParamsDO.jsonFilterText); 
 groupId = aesUtil.decrypt(groupId); 
 LogUtils.logMessage(Level.INFO,"SavingsLoanDocuments sanction_letter:= line 7 " + groupId); 

 String branchName = Utils.getJsonProperty("branchName", requestParamsDO.jsonFilterText); 
 branchName = aesUtil.decrypt(branchName); 
 LogUtils.logMessage(Level.INFO,"SavingsLoanDocuments sanction_letter:= line 8 " + branchName); 

 String regionName = Utils.getJsonProperty("regionName", requestParamsDO.jsonFilterText); 
 regionName = aesUtil.decrypt(regionName); 
 LogUtils.logMessage(Level.INFO,"SavingsLoanDocuments sanction_letter:= line 9 " + regionName); 

 String date = Utils.getJsonProperty("proposedDisbursementDate", requestParamsDO.jsonFilterText); 
 date = aesUtil.decrypt(date); 
 LogUtils.logMessage(Level.INFO,"SavingsLoanDocuments sanction_letter:= line 10 " + date); 

 String centerName = ReportUtils.getCenterName(groupId); 
 LogUtils.logMessage(Level.INFO,"SavingsLoanDocuments sanction_letter:= line 11 " + centerName); 

 // String date = 
 // Utils.getJsonProperty("proposedDisbursementDate",requestParamsDO.jsonFilterText); 
 String customerId = Utils.getJsonProperty("customerId", requestParamsDO.jsonFilterText); 
 customerId = aesUtil.decrypt(customerId); 
 LogUtils.logMessage(Level.INFO,"SavingsLoanDocuments sanction_letter:= line 12 " + customerId); 

 String customerIds = customerId.substring(0, customerId.length() - 1); 
 LogUtils.logMessage(Level.INFO,"SavingsLoanDocuments sanction_letter:= line 13 " + customerIds); 

 // String centerName = 
 // ReportUtils.getCenterName(Utils.getJsonProperty("groupId", 
 // requestParamsDO.jsonFilterText)); 
 String fontFile = Utils.getJsonProperty("fontFile", requestParamsDO.jsonFilterText); 
 LogUtils.logMessage(Level.INFO,"SavingsLoanDocuments sanction_letter:= line 14 " + fontFile); 

 String vcLanguages = Utils.getJsonProperty("vcLanguages", requestParamsDO.jsonFilterText); 
 vcLanguages = aesUtil.decrypt(vcLanguages); 
 LogUtils.logMessage(Level.INFO,"SavingsLoanDocuments sanction_letter:= line 15 " + vcLanguages); 
 // Set language context for sanction table header (thread-safe) 
 ReportUtils.setSanctionLanguage(vcLanguages); 
 String groupNumber1 = Utils.getJsonProperty("groupNumber1", requestParamsDO.jsonFilterText); 
 groupNumber1 = aesUtil.decrypt(groupNumber1); 
 String centerNumber1 = Utils.getJsonProperty("centerNumber1", requestParamsDO.jsonFilterText); 
 centerNumber1 = aesUtil.decrypt(centerNumber1); 
 String branchId = Utils.getJsonProperty("branchId", requestParamsDO.jsonFilterText); 
 branchId = aesUtil.decrypt(branchId); 
 String regionId = Utils.getJsonProperty("regionId", requestParamsDO.jsonFilterText); 
 regionId = aesUtil.decrypt(regionId); 

 String snactionData = ReportUtils.generateSanctionLetter1( 
 "call sp_rpt_savings_sanction_letter_document(\"" + customerIds + "\",\"" + Utils.quotedString(date) + "\")"); 
 LogUtils.logMessage(Level.INFO,"SavingsLoanDocuments sanction_letter:= line 15 " + snactionData); 

 // employerDetailsData = ReportUtils.sanctionLetterEmployeeDetails(groupId); 
 // LogUtils.logMessage(Level.INFO,"SavingsLoanDocuments sanction_letter Employee Details:= 
 // line 1 " + employerDetailsData); 

 // Load Template Based On Language 
 String templateFile = "Sanction.txt"; 

 if ("Hindi".equalsIgnoreCase(vcLanguages)){ 
 templateFile = "Sanction_Hindi.txt"; } 
 else if ("Gujarati".equalsIgnoreCase(vcLanguages)) { 
 templateFile = "Sanction_Gujarati.txt"; } 
 else if ("Kannada".equalsIgnoreCase(vcLanguages)) { 
 templateFile = "Sanction_Kannada.txt"; } 
 else if ("Marathi".equalsIgnoreCase(vcLanguages)) { 
 templateFile = "Sanction_Marathi.txt"; } 
 else if ("Tamil".equalsIgnoreCase(vcLanguages)) { 
 templateFile = "Sanction_Tamil.txt"; } 
 else if ("Telugu".equalsIgnoreCase(vcLanguages)) { 
 templateFile = "Sanction_Telugu.txt"; } 
 else if ("Odia".equalsIgnoreCase(vcLanguages)) { 
 templateFile = "Sanction_Odia.txt"; } 
 else{ 
 templateFile = "Sanction.txt"; 
 } 

 String htmlString = getReportHtmlString(filePath + templateFile); 
 LogUtils.logMessage(Level.INFO,"SavingsLoanDocuments sanction_letter:= line 16 " + filePath); 

 htmlString = htmlString.replace("#@#EmpName&Signature#@#", ""); 
 // LogUtils.logMessage(Level.INFO,"SavingsLoanDocuments sanction_letter Employee Details 
 // EmpName:= line 2 " + employerDetailsData[0]); 
 htmlString = htmlString.replace("#@#EmployeeCode#@#", ""); 
 // LogUtils.logMessage(Level.INFO,"SavingsLoanDocuments sanction_letter Employee Details 
 // EmpCode:= line 3 " + employerDetailsData[1]); 
 htmlString = htmlString.replace("#@#EmployeeDesignation#@#", "Branch Manager / Transaction Office"); 
 // LogUtils.logMessage(Level.INFO,"SavingsLoanDocuments sanction_letter Employee Details 
 // EmpDesignation:= line 4 " + employerDetailsData[2]); 

 htmlString = htmlString.replace("#@#groupName#@#", groupName + "(" + groupNumber1 + ")"); 
 LogUtils.logMessage(Level.INFO,"SavingsLoanDocuments sanction_letter:= line 17 " + groupName); 

 htmlString = htmlString.replace("#@#centerName#@#", centerName + "(" + centerNumber1 + ")"); 
 LogUtils.logMessage(Level.INFO,"SavingsLoanDocuments sanction_letter:= line 18 " + centerName); 

 htmlString = htmlString.replace("#@#DSCName#@#", branchName + "|" + branchId); 
 LogUtils.logMessage(Level.INFO,"SavingsLoanDocuments sanction_letter:= line 19 " + branchName); 

 htmlString = htmlString.replace("#@#RegionName#@#", regionName + "|" + regionId); 
 LogUtils.logMessage(Level.INFO,"SavingsLoanDocuments sanction_letter:= line 20 " + regionName); 

 htmlString = htmlString.replace("#@#LoanApplicationDate#@#", date); 
 LogUtils.logMessage(Level.INFO,"SavingsLoanDocuments sanction_letter:= line 21 " + date); 

 htmlString = htmlString.replace("#@#Place#@#", regionName); 
 LogUtils.logMessage(Level.INFO,"SavingsLoanDocuments sanction_letter:= line 22 " + regionName); 

 htmlString = htmlString.replace("#@#DATE#@#", date); 
 LogUtils.logMessage(Level.INFO,"SavingsLoanDocuments sanction_letter:= line 23 " + date); 

 htmlString = htmlString.replace("#@#tableData#@#", snactionData); 
 LogUtils.logMessage(Level.INFO,"SavingsLoanDocuments sanction_letter:= line 24 " + snactionData); 

 htmlString = htmlString.replace("#@#Version#@#", version); 
 LogUtils.logMessage(Level.INFO,"SavingsLoanDocuments sanction_letter:= line 25 " + version); 

 htmlString = htmlString.replace("#@#GroupIntegration#@#", integrationStatus.equals("FINISHED") ? " AD " : " BD "); 
 LogUtils.logMessage(Level.INFO,"SavingsLoanDocuments sanction_letter:= line 26 "); 

 LogUtils.logMessage(Level.INFO,"SavingsLoanDocuments sanction_letter:= line 27 inset logs start "); 
 bllSavingsLog.insert_savings_sanction_data(requestParamsDO); 
 LogUtils.logMessage(Level.INFO,"SavingsLoanDocuments sanction_letter:= line 28 inset logs end "); 

 clearOldReportsData(); 
 result1.put("reportId", Utils.getJsonProperty("reportId", requestParamsDO.jsonFilterText)); 
 result1.put("fileName", htmlString); 
 result.setResultObject(result1.toString()); 
 LogUtils.logMessage(Level.INFO,"SavingsLoanDocuments sanction_letter:= line 29 "); 

 } 
 else 
 { 
 WebLOSUtils.getErrorObject(result, "1011", "Invalid Session"); 
 } 
 } 
 catch (NullPointerException npe) 
 { 
 npe = new NullPointerException("Some data of this report found null values, please contact administrator."); 
 WebLOSUtils.getErrorObject(result, "1012", npe.getMessage()); 
 } 
 catch (Exception ex) 
 { 
 WebLOSUtils.getErrorObject(result, "1012", ex.getMessage()); 
 } 
 finally 
 { 
 ReportUtils.clearSanctionLanguage(); 
 } 

 return Response.status(200).entity(result).build(); 
 } 
@POST 
 @Path("/get_sanction_letter_vernacular_languages") 
 @Consumes(MediaType.APPLICATION_JSON) 
 public Response fetchSanctionLetterVernacularLanguages( 
 RequestParamsDO requestParamsDO) 
 { 
 RequestStatusDO result = new RequestStatusDO(); 

 try 
 { 
 if (SessionUtils.validateRequest(requestParamsDO, request)) 
 { 
 result.setResultObject(bllSavingsLoanDocuments.fetchSanctionLetterVernacularLanguages(requestParamsDO)); 
 } 
 else 
 { 
 WebLOSUtils.getErrorObject(result,"1011","Invalid Session"); 
 } 
 } 
 catch(Exception ex) 
 { 
 WebLOSUtils.getErrorObject(result,"1013",ex.getMessage()); 
 } 
 return Response.status(200).entity(result).build(); 
 } 

}package com.weblos.model; 

import org.json.JSONObject; 

import com.volksoftech.rest.AesUtil; 

public class RequestParamsDO 
{ 

 public String accessCode; 
 public String sessionId; 
 public long timeStamp; 
 public int mobileFlag; 
 public JSONObject jsonFilterText; 
 public String branchId; 
 public String routeCode; 
 public String rowguid; 
 public String productCategoryId; 
 public String hierarchyId; 
 public String hierarchyType; 
 public String designation; 
 AesUtil aesUtil = new AesUtil(128, 1000); 

 public int getMobileFlag() 
 { 
 return mobileFlag; 
 } 

 public void setMobileFlag(int mobileFlag) 
 { 
 this.mobileFlag = mobileFlag; 
 } 

 public String getProductCategoryId() 
 { 
 return productCategoryId; 
 } 

 public void setProductCategoryId(String productCategoryId) 
 { 
 this.productCategoryId = productCategoryId; 
 } 

 public String getRowguid() 
 { 
 return rowguid; 
 } 

 public void setRowguid(String rowguid) 
 { 
 this.rowguid = rowguid; 
 } 

 public JSONObject getJsonFilterText() 
 { 
 return jsonFilterText; 
 } 

 public void setJsonFilterText(JSONObject jsonFilterText) 
 { 
 this.jsonFilterText = jsonFilterText; 
 } 

 public String getAccessCode() 
 { 
 return accessCode; 
 } 

 public void setAccessCode(String accessCode) 
 { 
 if (accessCode != null && !accessCode.matches("\\d{6,9}")) 
 { 
 try 
 { 
 aesUtil.decrypt(accessCode); 
 } 
 catch (Exception e) 
 { 
 throw new IllegalArgumentException("accessCode must be a number inbetween 6 to 9 digits"); 
 } 

 } 
 this.accessCode = accessCode; 
 } 

 public String getBranchId() 
 { 
 return branchId; 
 } 

 public void setBranchId(String branchId) 
 { 
 this.branchId = branchId; 
 } 

 public String getRouteCode() 
 { 
 return routeCode; 
 } 

 public void setRouteCode(String routeCode) 
 { 
 this.routeCode = routeCode; 
 } 

 public String getSessionId() 
 { 
 return sessionId; 
 } 

 public void setSessionId(String sessionId) 
 { 
 this.sessionId = sessionId; 
 } 

 public long getTimeStamp() 
 { 
 return timeStamp; 
 } 

 public void setTimeStamp(long timeStamp) 
 { 
 this.timeStamp = timeStamp; 
 } 

 public String getHierarchyId() 
 { 
 return hierarchyId; 
 } 

 public void setHierarchyId(String hierarchyId) 
 { 
 this.hierarchyId = hierarchyId; 
 } 

 public String getHierarchyType() 
 { 
 return hierarchyType; 
 } 

 public void setHierarchyType(String hierarchyType) 
 { 
 this.hierarchyType = hierarchyType; 
 } 

 public String getDesignation() 
 { 
 return designation; 
 } 

 public void setDesignation(String designation) 
 { 
 this.designation = designation; 
 } 
} 

package com.volksoftech.reporting; 

import java.io.File; 
import java.io.FileOutputStream; 
import java.sql.Connection; 
import java.sql.PreparedStatement; 
import java.sql.ResultSet; 
import java.sql.ResultSetMetaData; 
import java.sql.Statement; 
import java.sql.Timestamp; 
import java.text.SimpleDateFormat; 
import java.util.ArrayList; 
import java.util.Arrays; 
import java.util.Date; 

import org.apache.http.NameValuePair; 
import org.apache.poi.hssf.util.HSSFColor; 
import org.apache.poi.ss.usermodel.BorderStyle; 
import org.apache.poi.xssf.usermodel.XSSFCell; 
import org.apache.poi.xssf.usermodel.XSSFCellStyle; 
import org.apache.poi.xssf.usermodel.XSSFCreationHelper; 
import org.apache.poi.xssf.usermodel.XSSFFont; 
import org.apache.poi.xssf.usermodel.XSSFRow; 
import org.apache.poi.xssf.usermodel.XSSFSheet; 
import org.apache.poi.xssf.usermodel.XSSFWorkbook; 

import com.volksoftech.common.Constants; 
import com.volksoftech.common.LogUtils;import java.util.logging.Level; 
import com.volksoftech.common.Utils; 
import com.volksoftech.dal.DalHelper; 
import com.volksoftech.dal.DalUtils; 
import com.weblos.model.BMHVChangesDO; 
import com.weblos.model.BMHVChangesFieldsDO; 
import com.weblos.model.CodeValueMasterDO; 
import com.weblos.model.ConfigValuesDO; 
import com.weblos.model.LoanProductEnrollmentDO; 
import com.weblos.model.MenuDO; 
import com.weblos.model.RepaymentSchduleDO; 
import com.weblos.model.RepaymentSchduleDatesDO; 
import com.weblos.model.RequestParamsDO; 
import com.volksoftech.common.LogUtils;import java.util.logging.Level; 

public class ReportUtils 
{ 
 private static final ThreadLocal<String> SANCTION_LANG = ThreadLocal.withInitial(() -> "English"); 
 public static String generateExcelReport(String sql, String reportName, String fileName, boolean autoColumnNames, NameValuePair... params) 
 throws Exception 
 { 
 String result = fileName; // file name 
 XSSFWorkbook workbook = new XSSFWorkbook(); 
 XSSFSheet worksheet = workbook.createSheet(reportName); 
 DalHelper dalHelper = new DalHelper(); 
 Connection connection = null; 
 Statement stmt = null; 
 ResultSet resultSet = null; 
 int rowNumber = 0; 
 short columnNumber = 0; 

 SimpleDateFormat simpleDateFormat = new SimpleDateFormat(Constants.DATE_TIME_FORMAT_LONG); 

 XSSFRow row = worksheet.createRow(rowNumber++); 
 columnNumber = 0; 
 rowNumber++; 
 XSSFCell cell = row.createCell(columnNumber++); 
 cell.setCellValue(reportName); 
 XSSFCellStyle cellStyle = workbook.createCellStyle(); 
 XSSFFont font = workbook.createFont(); 
 font.setBoldweight(XSSFFont.BOLDWEIGHT_BOLD); 
 font.setFontHeightInPoints((short) 14); 
 cellStyle.setFont(font); 
 cell.setCellStyle(cellStyle); 
 worksheet.setColumnWidth(0, worksheet.getColumnWidth(0) * 2); 
 columnNumber = 4; 
 cell = row.createCell(columnNumber++); 
 cell.setCellValue("Generated by Web LOS, " + simpleDateFormat.format(new Date())); 
 // cell.setCellStyle(cellStyle); 

 for (NameValuePair pair : params) 
 { 
 row = worksheet.createRow(rowNumber++); 
 columnNumber = 0; 

 cell = row.createCell(columnNumber++); 
 cell.setCellValue(pair.getName()); 
 cellStyle = workbook.createCellStyle(); 
 cellStyle.setFillForegroundColor(HSSFColor.GREY_25_PERCENT.index); 
 cellStyle.setFillPattern(XSSFCellStyle.SOLID_FOREGROUND); 
 cell.setCellStyle(cellStyle); 

 cell = row.createCell(columnNumber++); 
 cell.setCellValue(pair.getValue()); 
 cellStyle = workbook.createCellStyle(); 
 font = workbook.createFont(); 
 font.setBoldweight(XSSFFont.BOLDWEIGHT_BOLD); 
 cellStyle.setFont(font); 
 cell.setCellStyle(cellStyle); 
 } 
 try 
 { 
 connection = dalHelper.getConnection(); 
 stmt = connection.createStatement(); 
 stmt.setQueryTimeout(0); 
 resultSet = stmt.executeQuery(sql); 

 ResultSetMetaData resultSetMetaData1 = null; 
 if (resultSet != null) 
 { 
 resultSetMetaData1 = resultSet.getMetaData(); 
 int colCount = resultSetMetaData1.getColumnCount(); 
 row = worksheet.createRow(rowNumber++); 
 columnNumber = 0; 
 for (int i = 1; i <= colCount; i++) 
 { 
 cell = row.createCell(columnNumber++); 
 cell.setCellValue( 
 (autoColumnNames) ? makeDisplayLabelFromFieldName(resultSetMetaData1.getColumnName(i)) : resultSetMetaData1.getColumnName(i)); 
 cellStyle = workbook.createCellStyle(); 
 cellStyle.setFillForegroundColor(HSSFColor.GREY_40_PERCENT.index); 
 cellStyle.setFillPattern(XSSFCellStyle.SOLID_FOREGROUND); 
 cell.setCellStyle(cellStyle); 
 } 
 while (resultSet.next()) 
 { 
 row = worksheet.createRow(rowNumber++); 
 columnNumber = 0; 
 for (int i = 1; i <= colCount; i++) 
 { 
 cell = row.createCell(columnNumber++); 
 cell.setCellValue(resultSet.getString(i)); 
 int columnType = resultSetMetaData1.getColumnType(i); 
 LogUtils.logMessage(Level.INFO,"Column Name : " + resultSetMetaData1.getColumnName(i) + ", columnType : " + columnType); 
 switch (columnType) 
 { 
 case java.sql.Types.INTEGER: 
 cell.setCellValue(resultSet.getInt(i)); 
 break; 
 case java.sql.Types.SMALLINT: 
 cell.setCellValue(resultSet.getInt(i)); 
 break; 
 case java.sql.Types.TINYINT: 
 cell.setCellValue(resultSet.getInt(i)); 
 break; 
 case java.sql.Types.VARCHAR: 
 String strVal = resultSet.getString(i); 
 cell.setCellValue(strVal != null ? strVal : ""); 
 break; 
 case java.sql.Types.CHAR: 
 String strVal2 = resultSet.getString(i); 
 cell.setCellValue(strVal2 != null ? strVal2 : ""); 
 break; 
 case java.sql.Types.LONGVARCHAR: 
 String strVal1 = resultSet.getString(i); 
 cell.setCellValue(strVal1 != null ? strVal1 : ""); 
 break; 
 case java.sql.Types.TIMESTAMP: 
 Timestamp tsVal = resultSet.getTimestamp(i); 
 if(tsVal != null) { 
 java.util.Date dateVal = new Date(tsVal.getTime()); 
 cell.setCellValue(dateVal); 
 XSSFCellStyle dateTimeStyle = workbook.createCellStyle(); 
 XSSFCreationHelper createHelper = workbook.getCreationHelper(); 
 dateTimeStyle.setDataFormat(createHelper.createDataFormat().getFormat("dd-MM-yyyy HH:mm:ss")); 
 cell.setCellStyle(dateTimeStyle); 
 }else { 
 cell.setCellValue(""); 
 } 
 break; 
 default: 
 String defVal = resultSet.getString(i); 
 cell.setCellValue(defVal != null ? defVal : ""); 
 } 
 } 

 } 
 // row = worksheet.createRow(rowNumber++); 
 } 

 } 
 catch (Exception ex) 
 { 
 Utils.handleServerException("ReportUtils", "generateExcelReport", ex.getMessage(), ex); 
 throw ex; 
 } 
 finally 
 { 
 try 
 { 
 if (resultSet != null) 
 { 
 dalHelper.closeResultSet(resultSet); 
 } 
 // resultSet.close(); 
 } 
 catch (Exception ex1) 
 { 
 Utils.handleServerException("ReportUtils", "resultSet close", ex1.getMessage(), ex1); 
 } 
 finally 
 { 
 dalHelper.closeConnection(); 
 } 
 } 
 // Write the workbook in file system 
 FileOutputStream outputStream = null; 
 try 
 { 
 File file = new File(result); 
 System.out.println("generateExcelReport :" + file.getAbsolutePath()); 
 outputStream = new FileOutputStream(file); 
 workbook.write(outputStream); 
 workbook.close(); 
 } 
 catch (Exception ex) 
 { 
 Utils.handleServerException("ReportUtils", "generateExcelReport", ex.getMessage(), ex); 
 } 
 finally 
 { 
 try 
 { 
 if (outputStream != null) 
 { 
 outputStream.close(); 
 } 
 } 
 catch (Exception ex1) 
 { 
 Utils.handleServerException("ReportUtils", "generateExcelReport close", ex1.getMessage(), ex1); 
 } 
 } 

 return result; 
 } public static String generateExcelReportNew(String sql, String reportName, String fileName, boolean autoColumnNames, 
 NameValuePair... params) 
 { 
 String result = fileName; // file name 
 XSSFWorkbook workbook = new XSSFWorkbook(); 
 XSSFSheet worksheet = workbook.createSheet(reportName); 
 DalHelper dalHelper = new DalHelper(); 
 Connection connection = null; 
 Statement stmt = null; 
 ResultSet resultSet = null; 
 int rowNumber = 0; 
 int columnNumber = 0; 

 SimpleDateFormat simpleDateFormat = new SimpleDateFormat(Constants.DATE_TIME_FORMAT_LONG); 

 XSSFRow row = worksheet.createRow(rowNumber); 
 columnNumber = 0; 
 // rowNumber++; 
 XSSFCell cell = row.createCell(columnNumber); 
 // cell.setCellValue(reportName); 
 XSSFCellStyle cellStyle = workbook.createCellStyle(); 
 XSSFFont font = workbook.createFont(); 
 font.setBoldweight(XSSFFont.BOLDWEIGHT_BOLD); 
 font.setFontHeightInPoints((short) 14); 
 cellStyle.setFont(font); 
 // cell.setCellStyle(cellStyle); 
 worksheet.setColumnWidth(0, worksheet.getColumnWidth(0) * 2); 
 columnNumber = 4; 
 // cell = row.createCell(columnNumber++); 
 // cell.setCellValue("Generated by Web LOS, " + simpleDateFormat.format(new 
 // Date())); 
 // cell.setCellStyle(cellStyle); 

 for (NameValuePair pair : params) 
 { 
 row = worksheet.createRow(rowNumber); 
 columnNumber = 0; 

 cell = row.createCell(columnNumber++); 
 cell.setCellValue(pair.getName()); 
 cellStyle = workbook.createCellStyle(); 
 cellStyle.setFillForegroundColor(HSSFColor.GREY_25_PERCENT.index); 
 cellStyle.setFillPattern(XSSFCellStyle.SOLID_FOREGROUND); 
 cell.setCellStyle(cellStyle); 

 cell = row.createCell(columnNumber++); 
 cell.setCellValue(pair.getValue()); 
 cellStyle = workbook.createCellStyle(); 
 font = workbook.createFont(); 
 font.setBoldweight(XSSFFont.BOLDWEIGHT_BOLD); 
 cellStyle.setFont(font); 
 cell.setCellStyle(cellStyle); 
 } 
 try 
 { 
 connection = dalHelper.getConnection(); 
 stmt = connection.createStatement(); 
 resultSet = stmt.executeQuery(sql); 

 ResultSetMetaData resultSetMetaData1 = null; 
 if (resultSet != null) 
 { 
 resultSetMetaData1 = resultSet.getMetaData(); 
 int colCount = resultSetMetaData1.getColumnCount(); 
 row = worksheet.createRow(rowNumber++); 
 columnNumber = 0; 
 for (int i = 1; i <= colCount; i++) 
 { 
 cell = row.createCell(columnNumber++); 
 cell.setCellValue( 
 (autoColumnNames) ? makeDisplayLabelFromFieldName(resultSetMetaData1.getColumnName(i)) : resultSetMetaData1.getColumnName(i)); 
 cellStyle = workbook.createCellStyle(); 
 cellStyle.setFillForegroundColor(HSSFColor.GREY_40_PERCENT.index); 
 cellStyle.setFillPattern(XSSFCellStyle.SOLID_FOREGROUND); 
 cell.setCellStyle(cellStyle); 
 } 
 while (resultSet.next()) 
 { 
 row = worksheet.createRow(rowNumber++); 
 columnNumber = 0; 
 for (int i = 1; i <= colCount; i++) 
 { 
 cell = row.createCell(columnNumber++); 
 cell.setCellValue(resultSet.getString(i)); 
 } 
 } 
 } 
 } 
 catch (Exception ex) 
 { 
 LogUtils.logMessage(Level.INFO,"ReportUtils : generateExcelReport :" + ex.getMessage()); 
 } 
 finally 
 { 
 try 
 { 
 if (resultSet != null) 
 { 
 dalHelper.closeResultSet(resultSet); 
 } 
 // resultSet.close(); 
 } 
 catch (Exception ex1) 
 { 
 // Utils.handleServerException("ReportUtils", "resultSet close", 
 // ex1.getMessage(), ex1); 
 LogUtils.logMessage(Level.INFO,"ReportUtils : resultSet close :" + ex1.getMessage()); 
 } 
 finally 
 { 
 dalHelper.closeConnection(); 
 } 
 } 
 // Write the workbook in file system 
 FileOutputStream outputStream = null; 
 try 
 { 
 File file = new File(result); 
 // System.out.println("generateExcelReport :" + file.getAbsolutePath()); 
 LogUtils.logMessage(Level.INFO,"ReportUtils - generateExcelReportNew - File AbsolutePath = " + file.getAbsolutePath()); 
 outputStream = new FileOutputStream(file); 
 workbook.write(outputStream); 
 workbook.close(); 
 } 
 catch (Exception ex) 
 { 
 // Utils.handleServerException("ReportUtils", "generateExcelReport", 
 // ex.getMessage(), ex); 
 LogUtils.logMessage(Level.INFO,"ReportUtils :generateExcelReport :" + ex.getMessage()); 

 } 
 finally 
 { 
 try 
 { 
 if (outputStream != null) 
 { 
 outputStream.close(); 
 } 
 } 
 catch (Exception ex1) 
 { 
 // Utils.handleServerException("ReportUtils", "generateExcelReport close", 
 // ex1.getMessage(), ex1); 

 LogUtils.logMessage(Level.INFO,"ReportUtils :generateExcelReport close :" + ex1.getMessage()); 
 } 
 } 

 return result; 
 } 

public static String generateSanctionLetter1(String sql) 
 { 
 DalHelper dalHelper = new DalHelper(); 
 Connection connection = null; 
 Statement stmt = null; 
 ResultSet resultSet = null; 
 String[] rowData = null; 
 String resultData = ""; 
 try 
 { 
 connection = dalHelper.getConnection(); 
 stmt = connection.createStatement(); 
 resultSet = stmt.executeQuery(sql); 
 ResultSetMetaData resultSetMetaData1 = null; 
 if (resultSet != null) 
 { 
 resultSetMetaData1 = resultSet.getMetaData(); 
 int colCount = resultSetMetaData1.getColumnCount(); 
 rowData = new String[colCount]; 
 for (int i = 1; i <= colCount; i++) 
 { 
 rowData[i - 1] = resultSetMetaData1.getColumnLabel(i); 
 } 
 // Localized header by vcLanguages (fallback to DB header if not matched) 
 String[] localizedHeader = getSanctionHeaderByLanguage(rowData); 
 resultData = "<table style=\"border:1px solid black; border-collapse:collapse;width:100%; table-layout:fixed; font-size: 10pt;word-break: break-word;\"> " 
 + addSanctionTableHeader(localizedHeader); // for header 
 while (resultSet.next()) 
 { 
 rowData = new String[colCount]; 
 for (int i = 1; i <= colCount; i++) 
 { 
 rowData[i - 1] = resultSet.getString(i); 
 } 
 resultData += addSanctionTableRow(rowData); 
 } 
 resultData += "</table>"; 
 } 
 } 
 catch (Exception ex1) 
 { 
 Utils.handleServerException("ReportUtils", "generateSanctionLetter1 close", ex1.getMessage(), ex1); 
 } 
 finally 
 { 
 if (resultSet != null) 
 { 
 dalHelper.closeResultSet(resultSet); 
 } 
 } 
 return resultData; 
 } 
 private static String[] getSanctionHeaderByLanguage(String[] fallbackHeaders) 
 { 
 String lang = getSanctionLanguage(); 

 if ("Hindi".equalsIgnoreCase(lang)) { 
 return new String[] { 
 "AZw.H«$.", "H$O©Xma H$m Zm_/ n{V H$m Zm_", "Am`w", "dmoQ>a AmB©S>r", 
 "Amdo{XV bmoZ am{e", "ñdrH¥$V bmoZ am{e", "bmoZ H$s Ad{Y", "ã`mO H$s Xa", 
 "bmoZ H$m à`moOZ/ bmoZ à`moOZ H$s Cn-loUr", "J«mhH$ H$m hñVmja/ g§X^© H«$." 
 }; 
 } else if ("Gujarati".equalsIgnoreCase(lang)) { 
 return new String[] { 
 "A“y¾$d", "F>Zv$pf“y„ “pd/ ‘rs“y„ “pd", "Jdf", "dsv$pf Ap¡mM‘Ó", 
 "AfÆ L$f¡gu gp¡““u fL$d", "d„S|>f L$f¡gu gp¡““u fL$d", "gp¡““u dyv$s", "ìepS> v$f", 
 "gp¡““p¡ l¡sy/ gp¡““p¡ l¡sy ‘¡V$p- î¡Zu", "N°plL$“u klu/ k„v$c® “„." 
 }; 
 } else if ("Kannada".equalsIgnoreCase(lang)) { 
 return new String[] { 
 "PÜÅ™. ÓÜí.", "ÓÝÆWÝÃÜÃÜ ÖæÓÜÃÜá/±Ü£¿á ÖæÓÜÃÜá", "ÊÜ¿áÓÜáÕ", "ÊæäàoÃ… Iw", 
 "AiìÓÜÈÉÔ¨Ü ÓÝÆ¨Ü ÊæãñÜ¤", "ÊÜáígãÃÜáÊÜÞw¨Ü ÓÝÆ¨Ü ÊæãñÜ¤", "ÓÝÆ¨Ü AÊÜ—", "Ÿwx¿á¨ÜÃÜ", 
 "ÓÝÆ¨Ü E¨æªàÍÜ/ÓÝÆ¨Ü E¨æªàÍÜ¨Ü E±Ü&ÊÜWÜì", "WÝÅÖÜPÜÃÜ ÓÜ×/EÇæÉàS ÓÜí." 
 }; 
 } else if ("Marathi".equalsIgnoreCase(lang)) { 
 return new String[] { 
 "AZwH«$_m§H$", "H$O©XmamMo Zmd / nVrMo Zmd", "d`", "VXma AmoiInÍm", 
 "AO© Ho$cocr H$Om©Mr aŠH$_", "_§Oya Ho$cocr H$Om©Mr aŠH$_", "H$Om©Mm H$mcmdYr", "ì`mOXa", 
 "H$Om©Mm CÔoe / H$Om©Mm CÔoe Cn-àH$ma", "J«mhH$mMr ghr / g§X^© H«$_m§H$" 
 }; 
 } else if ("Tamil".equalsIgnoreCase(lang)) { 
 return new String[] { 
 "Á›øŒ Gs", "Phß ö£Ö£Á›ß ö£¯º / PnÁº ö£¯º", "Á¯x", "ÁõUPõÍº Aøh¯õÍ Amøh", 
 "@Põµ¨£mh Phß öuõøP", "AÝ©vUP¨£mh Phß öuõøP", "Phß Põ»®", "Ámi ÂQu®", 
 "Phß @|õUP® / Phß @|õUPzvß xøn ÁøP", "ÁõiUøP¯õÍº øPö¯õ¨£® / SÔ¨¦ Gs" 
 }; 
 } else if ("Telugu".equalsIgnoreCase(lang)) { 
 return new String[] { 
 "{Mæü.çÜ…QÅ", "Ææÿ$×ý{Væüïßý™èþ/¿æýÆæÿ¢ õ³Ææÿ$", "ÐèþÄæý$çÜ$Þ", "KrÆŠÿ Iyìþ", 
 "§æþÆæÿRêçÜ$¢ ^óþíÜ¯èþ ÌZ¯Œþ Æ>Õ", "Ðèþ$…lÆæÿ$ ^óþÄæý$ºyìþ¯èþ ÌZ¯Œþ Æ>Õ", "ÌZ¯Œþ ÐèþÅÐèþ«¨", "Ðèþyîþz Æóÿr$", 
 "ÌZ¯Œþ E§óþªÔèý…/ÌZ¯Œþ E§óþªÔèý… Eç³& MóürWÇ", "MæüçÜtÐèþ$ÆŠÿ çÜ…™èþMæü…/ÇçœÆðÿ¯ŒþÞ ¯èþ…." 
 }; 
 } else if ("Odia".equalsIgnoreCase(lang)) { 
 return new String[] { 
 "Lÿ÷.Óó", "J~S÷Üÿê†ÿæZÿ œÿæþ / Ó´æþêZÿ œÿæþ", "¯ÿßÓ", "{µÿæsÀÿ AæBÝ", 
 "¨÷{ßæS J~ ¨Àÿçþæ~", "þqëÀÿ J~ ¨Àÿçþæ~", "J~Àÿ ÓþßÓêþæ", "Óë™ ÜÿæÀÿ", 
 "J~ D{”É¿ / J~ D{”É¿ D¨-¯ÿSö", "S÷æÜÿLÿ Ó´æäÀÿ / Ó¢ÿµÿö ÓóQ¿æ" 
 }; 
 } 

 // Default English / DB header 
 return fallbackHeaders; 
 } 
 private static String getSanctionHeaderFont() { 

 String lang = getSanctionLanguage(); 

 if ("Hindi".equalsIgnoreCase(lang)){ return "face='SHREE-DEV-0714E Italic' class='hindi-text'";} 
 if ("Gujarati".equalsIgnoreCase(lang)) { return "face='SHREE-GUJ-0768E Bold' class='gujarati-text'";} 
 if ("Kannada".equalsIgnoreCase(lang)){ return "face='SHREE-KAN-0850 BoldItalic' class='kannada-text'";} 
 if ("Marathi".equalsIgnoreCase(lang)){ return "face='SHREE-DEV-0714E' class='marathi-text'";} 
 if ("Tamil".equalsIgnoreCase(lang)){ return "face='SHREE-TAM7-0802 Bold' class='tamil-text'";} 
 if ("Telugu".equalsIgnoreCase(lang)){ return "face='SHREE-TEL-0908 Italic' class='telugu-text'";} 
 if ("Odia".equalsIgnoreCase(lang)){ return "face='SHREE-ORI-0601E Bold' class='odia-text'";} 

 return "face='times, serif'"; 
 } 
 public static void setSanctionLanguage(String language) { 
 if (language == null || language.trim().isEmpty()) { 
 SANCTION_LANG.set("English"); 
 } else if (language.equalsIgnoreCase("Hindi")){ 
 SANCTION_LANG.set("Hindi"); 
 } else if (language.equalsIgnoreCase("Gujarati")) { 
 SANCTION_LANG.set("Gujarati"); 
 } else if (language.equalsIgnoreCase("Kannada")) { 
 SANCTION_LANG.set("Kannada"); 
 } else if (language.equalsIgnoreCase("Marathi")) { 
 SANCTION_LANG.set("Marathi"); 
 } else if (language.equalsIgnoreCase("Tamil")) { 
 SANCTION_LANG.set("Tamil"); 
 } else if (language.equalsIgnoreCase("Telugu")) { 
 SANCTION_LANG.set("Telugu"); 
 } else if (language.equalsIgnoreCase("Odia")) { 
 SANCTION_LANG.set("Odia"); 
 }else { 
 SANCTION_LANG.set("English"); 
 } 
 } 
 private static String getSanctionLanguage() { 
 String lang = SANCTION_LANG.get(); 
 return (lang == null || lang.trim().isEmpty()) ? "English" : lang.trim(); 
 } 
 public static void clearSanctionLanguage() { 
 SANCTION_LANG.remove(); 
 } 

} 
"BLL method in BllSavingsLoanDocuments" 
public ArrayList<LoanDocumentsDO>fetchSanctionLetterVernacularLanguages(RequestParamsDO requestParamsDO) throws Exception 
 { 
 return dalSavingsLoanDocuments.fetchSanctionLetterVernacularLanguages(requestParamsDO); 

 } 
 "DAL method in DalSavingsLoanDocuments" 
 public ArrayList<LoanDocumentsDO> fetchSanctionLetterVernacularLanguages(RequestParamsDO requestParamsDO) throws Exception 
 { 
 ArrayList<LoanDocumentsDO> result = new ArrayList<LoanDocumentsDO>(); 
 AesUtil aesUtil = new AesUtil(128, 1000); 
 Connection connection =null; 
 Statement stmt =null; 
 ResultSet resultSet =null; 

 String stateId = aesUtil.decrypt(Utils.getJsonProperty("stateId", requestParamsDO.jsonFilterText)); 
 String strsql = "select report_name,language,font_file,state_name,state_id from sanction_letter_vernacular_languages_fonts where delete_flag=0 and " 
 + "state_id = " + Utils.quotedString(stateId) + ""; 

 LoanDocumentsDO loanDocumentsDO =null; 

 try 
 { 
 connection =dalHelper.getConnection(); 
 stmt = connection.createStatement(); 
 resultSet = stmt.executeQuery(strsql); 

 while (resultSet.next()) 
 { 
 loanDocumentsDO = new LoanDocumentsDO(); 

 loanDocumentsDO.setVernacularReportName(resultSet.getString(1)); 
 loanDocumentsDO.setVernacularLanguages(resultSet.getString(2)); 
 loanDocumentsDO.setVernacularFontFile(resultSet.getString(3)); 
 loanDocumentsDO.setVernacularStateName(resultSet.getString(4)); 
 loanDocumentsDO.setVernacularStateId(resultSet.getString(5)); 

 result.add(loanDocumentsDO); 

 } 
 } 

 catch(Exception ex) 
 { 
 LogUtils.logMessage(Level.INFO," DalSavingsLoanDocuments - fetchSanctionLetterVernacularLanguages : " + ex.getMessage()); 
 throw ex; 
 } 

 finally 
 { 
 if(resultSet != null) 
 { 
 dalHelper.closeResultSet(resultSet); 
 } 
 } 

 return result; 
 } 
 public static String addSanctionTableHeader(String[] headerData) 
 { 
 String result = ""; 
 String headerWidth = ""; 
 String fontName = getSanctionHeaderFont(); // get font name and class name from, default font is : time, serif 
 result += "<tr>"; 
 for (int i = 0; i < headerData.length; i++) 
 { 
 headerWidth = tdWidthSanction(i); 
 result += "<th style=\"border:0.5px solid black;padding-left:2px;width:" + headerWidth 
 + "\" ><font " + fontName + " size=\"1\">"+ (headerData[i]) + " </font></th> "; 
 } 
 result += "</tr>"; 
 return result; 
 } 
 public static String addSanctionTableRow(String[] rowData) 
 { 
 String result = ""; 
 String cellWidth = ""; 
 result += "<tr style='height:30px;'>"; 
 for (int i = 0; i < rowData.length; i++) 
 { 
 cellWidth = tdWidthSanction(i); 
 result += addRowCell(rowData[i], cellWidth); 
 } 
 result += "</tr>"; 
 return result; 
 } 
 public static String tdWidthSanction(int columnNo) 
 { 
 String result = "15%"; 
 switch (columnNo) 
 { 
 case 0: 
 result = "3%"; 
 break; 
 case 1: 
 result = "9%"; 
 break; 
 case 2: 
 result = "3%"; 
 break; 
 case 3: 
 result = "5%"; 
 break; 
 case 4: 
 result = "6.5%"; 
 break; 
 case 5: 
 result = "6.5%"; 
 break; 
 case 6: 
 result = "3%"; 
 break; 
 case 7: 
 result = "3.5%"; 
 break; 
 case 8: 
 result = "9%"; 
 break; 
 case 9: 
 result = "9%"; 
 break; 

 } 
 return result; 
 } 
public static String addRowCell(String cellData, String cellWidth) 
 { 
 String result = ""; 
 result += "<td style=\"border:0.5px solid black;padding-left:2px;width:" + cellWidth + "\">"; 
 result += "<font face=\"times, serif\" size=\"1\"> " + cellData + "</font>"; 
 result += "</td>"; 
 return result; 
 }

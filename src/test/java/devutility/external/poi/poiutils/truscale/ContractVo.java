package devutility.external.poi.poiutils.truscale;

import devutility.external.poi.models.FieldColumnMap;

public class ContractVo {
	private String number;
	private String description;
	private String type;
	private String startDate;
	private String endDate;
	private String countryCode;
	private String currencyId;
	private String poNumber;
	private String state;
	private String billingFrequency;
	private String billingDay;
	private String buyingCustomerId;
	private String buyingCustomerContactId;
	private String billtoCustomerId;
	private String payerCustomerId;
	private String endCustomerId;
	private String endCustomerContactId;
	private String salesOffice;
	private String salesOrg;
	private String distributionChannel;
	private String division;
	private String paymentTerm;
	private String langugaeCode;
	private String geo;
	private String industry;

	public static FieldColumnMap<ContractVo> getFieldColumnMap() {
		FieldColumnMap<ContractVo> fieldColumnMap = new FieldColumnMap<>(ContractVo.class);
		fieldColumnMap.put(0, "number");
		fieldColumnMap.put(1, "description");
		fieldColumnMap.put(2, "type");
		fieldColumnMap.put(3, "startDate");
		fieldColumnMap.put(4, "endDate");
		fieldColumnMap.put(5, "countryCode");
		fieldColumnMap.put(6, "currencyId");
		fieldColumnMap.put(7, "poNumber");
		fieldColumnMap.put(8, "state");
		fieldColumnMap.put(9, "billingFrequency");
		fieldColumnMap.put(10, "billingDay");
		fieldColumnMap.put(11, "buyingCustomerId");
		fieldColumnMap.put(12, "buyingCustomerContactId");
		fieldColumnMap.put(13, "billtoCustomerId");
		fieldColumnMap.put(14, "payerCustomerId");
		fieldColumnMap.put(15, "endCustomerId");
		fieldColumnMap.put(16, "endCustomerContactId");
		return fieldColumnMap;
	}

	public String getNumber() {
		return number;
	}

	public void setNumber(String number) {
		this.number = number;
	}

	public String getDescription() {
		return description;
	}

	public void setDescription(String description) {
		this.description = description;
	}

	public String getType() {
		return type;
	}

	public void setType(String type) {
		this.type = type;
	}

	public String getStartDate() {
		return startDate;
	}

	public void setStartDate(String startDate) {
		this.startDate = startDate;
	}

	public String getEndDate() {
		return endDate;
	}

	public void setEndDate(String endDate) {
		this.endDate = endDate;
	}

	public String getCountryCode() {
		return countryCode;
	}

	public void setCountryCode(String countryCode) {
		this.countryCode = countryCode;
	}

	public String getCurrencyId() {
		return currencyId;
	}

	public void setCurrencyId(String currencyId) {
		this.currencyId = currencyId;
	}

	public String getPoNumber() {
		return poNumber;
	}

	public void setPoNumber(String poNumber) {
		this.poNumber = poNumber;
	}

	public String getState() {
		return state;
	}

	public void setState(String state) {
		this.state = state;
	}

	public String getBillingFrequency() {
		return billingFrequency;
	}

	public void setBillingFrequency(String billingFrequency) {
		this.billingFrequency = billingFrequency;
	}

	public String getBillingDay() {
		return billingDay;
	}

	public void setBillingDay(String billingDay) {
		this.billingDay = billingDay;
	}

	public String getBuyingCustomerId() {
		return buyingCustomerId;
	}

	public void setBuyingCustomerId(String buyingCustomerId) {
		this.buyingCustomerId = buyingCustomerId;
	}

	public String getBuyingCustomerContactId() {
		return buyingCustomerContactId;
	}

	public void setBuyingCustomerContactId(String buyingCustomerContactId) {
		this.buyingCustomerContactId = buyingCustomerContactId;
	}

	public String getBilltoCustomerId() {
		return billtoCustomerId;
	}

	public void setBilltoCustomerId(String billtoCustomerId) {
		this.billtoCustomerId = billtoCustomerId;
	}

	public String getPayerCustomerId() {
		return payerCustomerId;
	}

	public void setPayerCustomerId(String payerCustomerId) {
		this.payerCustomerId = payerCustomerId;
	}

	public String getEndCustomerId() {
		return endCustomerId;
	}

	public void setEndCustomerId(String endCustomerId) {
		this.endCustomerId = endCustomerId;
	}

	public String getEndCustomerContactId() {
		return endCustomerContactId;
	}

	public void setEndCustomerContactId(String endCustomerContactId) {
		this.endCustomerContactId = endCustomerContactId;
	}

	public String getSalesOffice() {
		return salesOffice;
	}

	public void setSalesOffice(String salesOffice) {
		this.salesOffice = salesOffice;
	}

	public String getSalesOrg() {
		return salesOrg;
	}

	public void setSalesOrg(String salesOrg) {
		this.salesOrg = salesOrg;
	}

	public String getDistributionChannel() {
		return distributionChannel;
	}

	public void setDistributionChannel(String distributionChannel) {
		this.distributionChannel = distributionChannel;
	}

	public String getDivision() {
		return division;
	}

	public void setDivision(String division) {
		this.division = division;
	}

	public String getPaymentTerm() {
		return paymentTerm;
	}

	public void setPaymentTerm(String paymentTerm) {
		this.paymentTerm = paymentTerm;
	}

	public String getLangugaeCode() {
		return langugaeCode;
	}

	public void setLangugaeCode(String langugaeCode) {
		this.langugaeCode = langugaeCode;
	}

	public String getGeo() {
		return geo;
	}

	public void setGeo(String geo) {
		this.geo = geo;
	}

	public String getIndustry() {
		return industry;
	}

	public void setIndustry(String industry) {
		this.industry = industry;
	}

	@Override
	public String toString() {
		return "TruscaleContractExcelVo [number=" + number + ", description=" + description + ", type=" + type + ", startDate=" + startDate + ", endDate=" + endDate + ", countryCode=" + countryCode + ", currencyId=" + currencyId + ", poNumber=" + poNumber
				+ ", state=" + state + ", billingFrequency=" + billingFrequency + ", billingDay=" + billingDay + ", buyingCustomerId=" + buyingCustomerId + ", buyingCustomerContactId=" + buyingCustomerContactId + ", billtoCustomerId=" + billtoCustomerId
				+ ", payerCustomerId=" + payerCustomerId + ", endCustomerId=" + endCustomerId + ", endCustomerContactId=" + endCustomerContactId + ", salesOffice=" + salesOffice + ", salesOrg=" + salesOrg + ", distributionChannel=" + distributionChannel
				+ ", division=" + division + ", paymentTerm=" + paymentTerm + ", langugaeCode=" + langugaeCode + ", geo=" + geo + ", industry=" + industry + "]";
	}
}
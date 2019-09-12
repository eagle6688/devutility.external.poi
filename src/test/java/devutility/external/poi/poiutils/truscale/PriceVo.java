package devutility.external.poi.poiutils.truscale;

public class PriceVo {
	private String customerId;
	private String billingConfigId;
	private String nodeStage;
	private String utilization;
	private String utlizationPrice;

	public PriceVo() {
	}

	public PriceVo(String customerId, String billingConfigId, String nodeStage, String utilization, String utlizationPrice) {
		this.customerId = customerId;
		this.billingConfigId = billingConfigId;
		this.nodeStage = nodeStage;
		this.utilization = utilization;
		this.utlizationPrice = utlizationPrice;
	}

	@Override
	public String toString() {
		return "TruscalePriceExcelVo{" + "customerId='" + customerId + '\'' + ", billingConfigId='" + billingConfigId + '\'' + ", nodeStage='" + nodeStage + '\'' + ", utilization='" + utilization + '\'' + ", utlizationPrice='" + utlizationPrice + '\'' + '}';
	}

	public String getCustomerId() {
		return customerId;
	}

	public void setCustomerId(String customerId) {
		this.customerId = customerId;
	}

	public String getBillingConfigId() {
		return billingConfigId;
	}

	public void setBillingConfigId(String billingConfigId) {
		this.billingConfigId = billingConfigId;
	}

	public String getNodeStage() {
		return nodeStage;
	}

	public void setNodeStage(String nodeStage) {
		this.nodeStage = nodeStage;
	}

	public String getUtilization() {
		return utilization;
	}

	public void setUtilization(String utilization) {
		this.utilization = utilization;
	}

	public String getUtlizationPrice() {
		return utlizationPrice;
	}

	public void setUtlizationPrice(String utlizationPrice) {
		this.utlizationPrice = utlizationPrice;
	}
}

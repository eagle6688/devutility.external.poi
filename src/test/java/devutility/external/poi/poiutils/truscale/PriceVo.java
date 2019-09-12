package devutility.external.poi.poiutils.truscale;

import java.math.BigDecimal;

public class PriceVo {
	private String customerId;
	private String billingConfigId;
	private String nodeStage;
	private String utilization;
	private String utlizationPrice;

	@Override
	public String toString() {
		try {
			System.out.println(utlizationPrice);
			BigDecimal bigDecimal = new BigDecimal(utlizationPrice);
			bigDecimal = bigDecimal.setScale(0, BigDecimal.ROUND_HALF_UP);
			return "TruscalePriceExcelVo{" + "customerId='" + customerId + '\'' + ", billingConfigId='" + billingConfigId + '\'' + ", nodeStage='" + nodeStage + '\'' + ", utilization='" + utilization + '\'' + ", utlizationPrice='" + bigDecimal.intValue()
					+ '\'' + '}';
		} catch (Exception e) {
			return "";
		}
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

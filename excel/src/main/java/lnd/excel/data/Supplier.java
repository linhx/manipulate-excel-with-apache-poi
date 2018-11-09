package lnd.excel.data;

/**
 * @author linhnguyendinh
 */
public class Supplier {
    Integer unitPrice;
    Integer totalAmount;
    String offer;
    String sampleSubmited;
    String remarks;

    public Integer getUnitPrice() {
        return unitPrice;
    }

    public void setUnitPrice(Integer unitPrice) {
        this.unitPrice = unitPrice;
    }

    public Integer getTotalAmount() {
        return totalAmount;
    }

    public void setTotalAmount(Integer totalAmount) {
        this.totalAmount = totalAmount;
    }

    public String getOffer() {
        return offer;
    }

    public void setOffer(String offer) {
        this.offer = offer;
    }

    public String getSampleSubmited() {
        return sampleSubmited;
    }

    public void setSampleSubmited(String sampleSubmited) {
        this.sampleSubmited = sampleSubmited;
    }

    public String getRemarks() {
        return remarks;
    }

    public void setRemarks(String remarks) {
        this.remarks = remarks;
    }
}

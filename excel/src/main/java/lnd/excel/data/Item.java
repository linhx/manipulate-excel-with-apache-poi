package lnd.excel.data;

/**
 * @author linhnguyendinh
 */
public class Item {
    String itemRef;
    String desc;
    Integer quatity;

    public String getItemRef() {
        return itemRef;
    }

    public void setItemRef(String itemRef) {
        this.itemRef = itemRef;
    }

    public String getDesc() {
        return desc;
    }

    public void setDesc(String desc) {
        this.desc = desc;
    }

    public Integer getQuatity() {
        return quatity;
    }

    public void setQuatity(Integer quatity) {
        this.quatity = quatity;
    }
}

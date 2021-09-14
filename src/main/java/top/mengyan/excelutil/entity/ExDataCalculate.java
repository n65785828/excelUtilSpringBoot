package top.mengyan.excelutil.entity;



import java.math.BigDecimal;


public class ExDataCalculate {
    private String code;
    private String name;
    private BigDecimal upV;//涨幅
    private BigDecimal upSpeed;//涨速
    private BigDecimal openP;//开盘
    private BigDecimal nowVolume;//现量
    private BigDecimal liuTongZ;//流通市值Z
    private BigDecimal totalMoney;//总金额
    private BigDecimal openMoney;//开盘金额
    private BigDecimal closeVar;//封单额
    private BigDecimal flowMarketVaR;//流通市值
    private BigDecimal changeHand;// 换手%
    private BigDecimal nowPrice;//现价

    public String getCode() {
        return code;
    }

    public void setCode(String code) {
        this.code = code;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public BigDecimal getUpV() {
        return upV;
    }

    public void setUpV(BigDecimal upV) {
        this.upV = upV;
    }

    public BigDecimal getUpSpeed() {
        return upSpeed;
    }

    public void setUpSpeed(BigDecimal upSpeed) {
        this.upSpeed = upSpeed;
    }

    public BigDecimal getOpenP() {
        return openP;
    }

    public void setOpenP(BigDecimal openP) {
        this.openP = openP;
    }

    public BigDecimal getNowVolume() {
        return nowVolume;
    }

    public void setNowVolume(BigDecimal nowVolume) {
        this.nowVolume = nowVolume;
    }

    public BigDecimal getLiuTongZ() {
        return liuTongZ;
    }

    public void setLiuTongZ(BigDecimal liuTongZ) {
        this.liuTongZ = liuTongZ;
    }

    public BigDecimal getTotalMoney() {
        return totalMoney;
    }

    public void setTotalMoney(BigDecimal totalMoney) {
        this.totalMoney = totalMoney;
    }

    public BigDecimal getOpenMoney() {
        return openMoney;
    }

    public void setOpenMoney(BigDecimal openMoney) {
        this.openMoney = openMoney;
    }

    public BigDecimal getCloseVar() {
        return closeVar;
    }

    public void setCloseVar(BigDecimal closeVar) {
        this.closeVar = closeVar;
    }

    public BigDecimal getFlowMarketVaR() {
        return flowMarketVaR;
    }

    public void setFlowMarketVaR(BigDecimal flowMarketVaR) {
        this.flowMarketVaR = flowMarketVaR;
    }

    public BigDecimal getChangeHand() {
        return changeHand;
    }

    public void setChangeHand(BigDecimal changeHand) {
        this.changeHand = changeHand;
    }

    public BigDecimal getNowPrice() {
        return nowPrice;
    }

    public void setNowPrice(BigDecimal nowPrice) {
        this.nowPrice = nowPrice;
    }
}

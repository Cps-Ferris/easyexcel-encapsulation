package cn.cps.easyexcel.test;

import com.alibaba.excel.annotation.ExcelProperty;

/**
 * @Author: Cai Peishen
 * @Date: 2020/12/2 14:15
 * @Description: EasyExcel填充
 */
public class FillInfo {

    //订单编号
    private String orderNo;

    //订单配款额
    private String peiMoney;

    //订单余额
    private String yuMoney;

    //商品名称
    private String goodsName;

    //合同金额
    private String heTongMoney;

    //合同量
    private String heTongCount;

    //卖方发货金额
    private String faHuoMoney;

    //卖方发货量
    private String faHuoCount;

    //收货结算金额
    private String jieSuanMoney;

    //结算量
    private String jieSuanCount;

    //订单状态
    private String status;

    public String getOrderNo() {
        return orderNo;
    }

    public void setOrderNo(String orderNo) {
        this.orderNo = orderNo;
    }

    public String getPeiMoney() {
        return peiMoney;
    }

    public void setPeiMoney(String peiMoney) {
        this.peiMoney = peiMoney;
    }

    public String getYuMoney() {
        return yuMoney;
    }

    public void setYuMoney(String yuMoney) {
        this.yuMoney = yuMoney;
    }

    public String getGoodsName() {
        return goodsName;
    }

    public void setGoodsName(String goodsName) {
        this.goodsName = goodsName;
    }

    public String getHeTongMoney() {
        return heTongMoney;
    }

    public void setHeTongMoney(String heTongMoney) {
        this.heTongMoney = heTongMoney;
    }

    public String getHeTongCount() {
        return heTongCount;
    }

    public void setHeTongCount(String heTongCount) {
        this.heTongCount = heTongCount;
    }

    public String getFaHuoMoney() {
        return faHuoMoney;
    }

    public void setFaHuoMoney(String faHuoMoney) {
        this.faHuoMoney = faHuoMoney;
    }

    public String getFaHuoCount() {
        return faHuoCount;
    }

    public void setFaHuoCount(String faHuoCount) {
        this.faHuoCount = faHuoCount;
    }

    public String getJieSuanMoney() {
        return jieSuanMoney;
    }

    public void setJieSuanMoney(String jieSuanMoney) {
        this.jieSuanMoney = jieSuanMoney;
    }

    public String getJieSuanCount() {
        return jieSuanCount;
    }

    public void setJieSuanCount(String jieSuanCount) {
        this.jieSuanCount = jieSuanCount;
    }

    public String getStatus() {
        return status;
    }

    public void setStatus(String status) {
        this.status = status;
    }
}

package com.zhuang.poi.model;

import com.zhuang.poi.excel.ExcelColumn;

/**
 * Created by zhuang on 12/30/2017.
 */

public class Product {

    @ExcelColumn(name = "名称")
    private String name;

    @ExcelColumn(name = "价格")
    private Float price;

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public Float getPrice() {
        return price;
    }

    public void setPrice(Float price) {
        this.price = price;
    }

    @Override
    public String toString() {
        return "Product [name=" + name + ", price=" + price + "]";
    }


}

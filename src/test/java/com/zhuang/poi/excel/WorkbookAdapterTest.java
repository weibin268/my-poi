package com.zhuang.poi.excel;

import com.zhuang.poi.model.AreaCodeInfo;
import com.zhuang.poi.model.Product;
import org.junit.Test;

import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.nio.file.StandardOpenOption;
import java.util.List;

/**
 * Created by zhuang on 12/30/2017.
 */
public class WorkbookAdapterTest {
    @Test
    public void toList() throws Exception {
        String pathname = this.getClass().getClassLoader().getResource("Product.xlsx").getPath();
        InputStream inputStream = new FileInputStream(new File(pathname));
        WorkbookAdapter workbookAdapter = new WorkbookAdapter(ExcelType.XLSX, inputStream);
        workbookAdapter.setSkipRowCount(0);
        List<Product> products = workbookAdapter.setRowAdaptHandler(context -> {
            Product product = (Product) context.getDataRow();
            if (product.getPrice() != null && product.getPrice() > 1) {
                return false;
            }
            return true;
        }).toList(Product.class);

        for (Product product : products) {
            System.out.println(product);
        }
    }

    @Test
    public void readAreaCodeInfo() throws Exception {
        String pathname = this.getClass().getClassLoader().getResource("AreaCodeInfo.xlsx").getPath();
        InputStream inputStream = new FileInputStream(new File(pathname));
        WorkbookAdapter workbookAdapter = new WorkbookAdapter(ExcelType.XLSX, inputStream);
        List<AreaCodeInfo> areaCodeInfoList = workbookAdapter.setRowAdaptHandler(new WorkbookAdapter.AdaptHandler() {
            public boolean handle(WorkbookAdapter.AdaptContext context) {
                return true;
            }
        }).toList(AreaCodeInfo.class);
        for (AreaCodeInfo areaCodeInfo : areaCodeInfoList) {
            String line = areaCodeInfo.getAreaCode() + "," + areaCodeInfo.getCityName() + "," + areaCodeInfo.getProvinceName() + "\n";

            Files.write(Paths.get("d:\\temp\\areacode.txt"), line.getBytes(), StandardOpenOption.APPEND, StandardOpenOption.CREATE);
        }
    }

}
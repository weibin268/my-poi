package com.zhuang.poi.model;

import com.zhuang.poi.excel.ExcelColumn;

/**
 * Created by zhuang on 12/30/2017.
 */

public class AreaCodeInfo {

	  @ExcelColumn(name = "序号")
	    private String seq;

	    @ExcelColumn(name = "省份")
	    private String provinceName;
	    
	    @ExcelColumn(name = "城")
	    private String cityName;
	    
	    @ExcelColumn(name = "区号")
	    private String areaCode;

		public String getSeq() {
			return seq;
		}

		public void setSeq(String seq) {
			this.seq = seq;
		}

		public String getProvinceName() {
			return provinceName;
		}

		public void setProvinceName(String provinceName) {
			this.provinceName = provinceName;
		}

		public String getCityName() {
			return cityName;
		}

		public void setCityName(String cityName) {
			this.cityName = cityName;
		}

		public String getAreaCode() {
			return areaCode;
		}

		public void setAreaCode(String areaCode) {
			this.areaCode = areaCode;
		}

		
		@Override
		public String toString() {
			return "Product [seq=" + seq + ", provinceName=" + provinceName + ", cityName=" + cityName + ", areaCode="
					+ areaCode + "]";
		}
	    

}

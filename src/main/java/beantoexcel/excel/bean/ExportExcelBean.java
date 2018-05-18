package beantoexcel.excel.bean;

import java.util.LinkedHashMap;
import java.util.List;

/**
 * @author guoxy
 * @description Excel简单通用工具
 * @create 2018-05-18 18:28
 **/
public class ExportExcelBean<T>{
	/**
	 * 要填充的内容
	 */
	private List<T> contentList;

	/**
	 * 表列列头名称
	 */
	private LinkedHashMap<String, String> keyMap;
	/**
	 * 分表名
	 */
	private String sheetName;

	public List<T> getContentList() {
		return contentList;
	}

	public void setContentList(List<T> contentList) {
		this.contentList = contentList;
	}

	public LinkedHashMap<String, String> getKeyMap() {
		return keyMap;
	}

	public void setKeyMap(LinkedHashMap<String, String> keyMap) {
		this.keyMap = keyMap;
	}

	public String getSheetName() {
		return sheetName;
	}

	public void setSheetName(String sheetName) {
		this.sheetName = sheetName;
	}

}

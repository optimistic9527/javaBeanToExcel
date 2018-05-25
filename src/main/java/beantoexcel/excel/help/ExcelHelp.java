package beantoexcel.excel.help;


import beantoexcel.excel.annotation.ExcelSheet;
import beantoexcel.excel.annotation.SheetCol;
import beantoexcel.excel.bean.ExportExcelBean;
import jxl.CellView;
import jxl.SheetSettings;
import jxl.Workbook;
import jxl.write.*;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.sql.Timestamp;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.stream.Collectors;

/**
 * @author guoxy
 * @description Excel简单通用工具
 * @create 2018-05-18 18:28
 **/
public class ExcelHelp<T> {
	//字体大小
	private int fontSize;
	//标题
	private String title;
	//时间格式
	private String timeFormat;

	private static final String DEFAULT_TIME_FORMAT = "yyyy-MM-dd HH:mm:ss";

	//用于存贮被注解修饰的bean的属性数据
	private HashMap<String, Field> fieldMap = new HashMap<>();

	// 标题（列头）样式
	private WritableCellFormat titleFormat;
	//列头样式
	private WritableCellFormat headFormat;
	// 正文样式
	private WritableCellFormat bodyCellFormat;
	//统计样式
	private WritableCellFormat sumCellFormat;

	private WritableWorkbook workbook;

	private ExcelHelp() {
		titleFormat = ShareFunction.setTitleCellFormat();
		headFormat = ShareFunction.setHeaderCellFormat();
		bodyCellFormat = ShareFunction.setBodyCellFormat();
		sumCellFormat = ShareFunction.setBodyRedCellFormat();
		timeFormat = DEFAULT_TIME_FORMAT;
		fontSize = 10;
	}

	private void export(ExportExcelBean<T> exportExcelBean, OutputStream os, Integer... cloNow) {
		try {
			workbook = Workbook.createWorkbook(os);
			addSheet(exportExcelBean.getKeyMap(), exportExcelBean.getContentList(), exportExcelBean.getSheetName(), cloNow);
			workbook.write();
			workbook.close();
			//关闭流
			os.flush();
			os.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	private void addSheet(LinkedHashMap<String, String> keyMap, List<T> listContent, String sheetName, Integer... cloNum)
			throws Exception{
		// 创建名为sheetName的工作表
		WritableSheet sheet = workbook.createSheet(sheetName, 0);
		// 设置标题,标题内容为keyMap中的value值
		//合并4个参数分别为开始列，开始行，合并到那一列，合并到哪一行
		int maxCloSize = keyMap.size();
		sheet.mergeCells(0, 0, maxCloSize-1, 0);
		sheet.addCell(new Label(0, 0, title, titleFormat));

		//冻结表头
		SheetSettings settings = sheet.getSettings();
		settings.setVerticalFreeze(2);

		//设置列头
		Iterator<Map.Entry<String, String>> headIterator = keyMap.entrySet().iterator();
		int titleIndex = 0;
		while (headIterator.hasNext()) {
			Map.Entry<String, String> entry = headIterator.next();
			sheet.addCell(new Label(titleIndex++, 1, entry.getValue(), headFormat));
		}
		Map<Integer, Double> cloCountMap = null;
		//过滤掉超过列数的数字，并映射成初始map
		if (cloNum != null) {
			cloCountMap = Arrays.stream(cloNum).filter(i->i<maxCloSize).collect(Collectors.toMap(i -> i, i -> 0.0));
		}
		// 设置正文内容
		for (int row = 0, size = listContent.size(); row < size; row++) {
			Iterator<Map.Entry<String, String>> headContent = keyMap.entrySet().iterator();
			int col = 0;
			while (headContent.hasNext()) {
				Map.Entry<String, String> entry = headContent.next();
				String key = entry.getKey();
				Field field = fieldMap.get(key);
				Object content = field.get(listContent.get(row));
				if (cloCountMap != null && cloCountMap.containsKey(col)) {
					try {
						cloCountMap.put(col, cloCountMap.get(col) + Double.parseDouble(String.valueOf(content)));
					} catch (Exception e) {
						//发生异常代表想要统计的列并不能转换成数字，
						cloCountMap.remove(col);
					}
				}
				Label label = getContentLabel(col, row + 2, field, content);
				col++;
				sheet.addCell(label);
			}
		}
		//统计map可用事，生成统计列
		if (cloCountMap != null&&!cloCountMap.isEmpty()) {
			sheet.addCell(new Label(0, listContent.size() + 2, "总计", sumCellFormat));
			for (Map.Entry<Integer, Double> integerIntegerEntry : cloCountMap.entrySet()) {
				sheet.addCell(new Label(integerIntegerEntry.getKey(), listContent.size() + 2, String.valueOf(integerIntegerEntry.getValue()), sumCellFormat));
			}
		}
		setAutoSize(sheet, maxCloSize, listContent.size());
	}

	public static ExcelHelp getExcelHelp() {
		return new ExcelHelp();
	}

	/**
	 *
	 * @param os 可以是响应流，可以是io流
	 * @param dataSource 需要生成excel的bean的列表
	 * @param title excel的标题是什么
	 * @param cloNow 需要去统计那写列的数据（列是从0开始的）
	 * @throws IOException
	 */
	public final void exportByAnnotation(OutputStream os, List<T> dataSource,String title,Integer... cloNow) {
		this.title = title;
		ExportExcelBean<T> sheetBeanByAnnotation = getSheetBeanByAnnotation(dataSource);
		export(sheetBeanByAnnotation, os, cloNow);
	}

	/**
	 *
	 * @param response http响应
	 * @param dataSource 需要生成excel的bean的列表
	 * @param title excel的标题是什么
	 * @param cloNow 需要去统计那写列的数据（列是从0开始的）
	 * @throws IOException
	 */
	public final void httpExport(HttpServletResponse response, List<T> dataSource, String title, Integer... cloNow) throws IOException {
		response.setContentType("application/vnd.ms-excel");
		response.setHeader("Content-Disposition", "attachment;filename=document.xls");
		exportByAnnotation(response.getOutputStream(), dataSource, title,cloNow);
	}

	//获取bean里面的注解数据
	private ExportExcelBean<T> getSheetBeanByAnnotation(List<T> sheet) {
		T row = sheet.get(0);
		Class<?> clazz = row.getClass();
		ExportExcelBean<T> sheetBean = new ExportExcelBean<>();
		sheetBean.setContentList(sheet);
		// 设置表名
		if (clazz.isAnnotationPresent(ExcelSheet.class)) {
			sheetBean.setSheetName(clazz.getAnnotation(ExcelSheet.class).name());
		} else {
			sheetBean.setSheetName("defaultSheet");
		}

		// 设置要展示的列
		LinkedHashMap<String, String> keyMap = new LinkedHashMap<>();
		Field[] fields = clazz.getDeclaredFields();
		for (Field field : fields) {
			field.setAccessible(true);
			if (field.isAnnotationPresent(SheetCol.class)) {
				String key = field.toString().substring(field.toString().lastIndexOf(".") + 1);
				keyMap.put(key, field.getAnnotation(SheetCol.class).value());
				fieldMap.put(key, field);
			}
		}
		sheetBean.setKeyMap(keyMap);
		return sheetBean;
	}

	/**
	 * 每个单元格的内容及格式
	 */
	protected Label getContentLabel(int col, int row, Field field, Object content) {
		String contentStr;
		contentStr = null != content ? content.toString() : "";
		// 如果是时间类型。要格式化成标准时间格式
		String timeStr = getTimeFormatValue(field, content);
		// timeStr不为空，说明是时间类型
		if (null != timeStr && !timeStr.trim().equals("")) {
			contentStr = timeStr;
		}
		return new Label(col, row, contentStr, bodyCellFormat);
	}

	/**
	 * 宽度自适应
	 */
	private void setAutoSize(WritableSheet sheet, int colNum, int rowNum) {
		for (int i = 0; i < colNum; i++) {
			int maxLength = 0;
			CellView cell = sheet.getColumnView(i);
			for (int j = 0; j < rowNum; j++) {
				maxLength = Math.max(sheet.getCell(i, j).getContents().getBytes().length, maxLength);
			}
			cell.setSize(25 * fontSize * maxLength);
			sheet.setColumnView(i, cell);
		}
	}

	/**
	 * 获取格式化后的时间串
	 */
	protected String getTimeFormatValue(Field field, Object content) {
		String timeFormatVal = "";
		if (field.getType().getName().equals(Timestamp.class.getName())) {
			Timestamp time = (Timestamp) content;
			timeFormatVal = longTimeTypeToStr(time.getTime(), timeFormat);
		} else if (field.getType().getName().equals(Date.class.getName())) {
			Date time = (Date) content;
			timeFormatVal = longTimeTypeToStr(time.getTime(), timeFormat);
		}

		return timeFormatVal;
	}

	/**
	 * 格式化时间
	 */
	protected String longTimeTypeToStr(long time, String formatType) {

		String strTime = "";
		if (time >= 0) {
			SimpleDateFormat sDateFormat = new SimpleDateFormat(formatType);
			strTime = sDateFormat.format(new Date(time));

		}
		return strTime;

	}


	public WritableCellFormat getHeadFormat() {
		return headFormat;
	}

	public void setHeadFormat(WritableCellFormat headFormat) {
		this.headFormat = headFormat;
	}

	public int getFontSize() {
		return fontSize;
	}

	public void setFontSize(int fontSize) {
		this.fontSize = fontSize;
	}

	public String getTitle() {
		return title;
	}

	public void setTitle(String title) {
		this.title = title;
	}

	public WritableCellFormat getTitleFormat() {
		return titleFormat;
	}

	public void setTitleFormat(WritableCellFormat titleFormat) {
		this.titleFormat = titleFormat;
	}

	public WritableCellFormat getBodyCellFormat() {
		return bodyCellFormat;
	}

	public void setBodyCellFormat(WritableCellFormat bodyCellFormat) {
		this.bodyCellFormat = bodyCellFormat;
	}

	public WritableCellFormat getSumCellFormat() {
		return sumCellFormat;
	}

	public void setSumCellFormat(WritableCellFormat sumCellFormat) {
		this.sumCellFormat = sumCellFormat;
	}
}

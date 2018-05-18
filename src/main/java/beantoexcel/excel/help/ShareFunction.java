package beantoexcel.excel.help;

import java.text.SimpleDateFormat;
import java.util.Date;

import jxl.format.Alignment;
import jxl.format.Border;
import jxl.format.BorderLineStyle;
import jxl.format.Colour;
import jxl.format.UnderlineStyle;
import jxl.format.VerticalAlignment;
import jxl.write.WritableCellFormat;
import jxl.write.WritableFont;


public class ShareFunction {

	//时间转换
	public static String setTimeFormat(Long lg, String fomart) {
		SimpleDateFormat sdf = new SimpleDateFormat(fomart);
		return sdf.format(new Date(lg));
	}

	//设置excel标题样式
	public static WritableCellFormat setTitleCellFormat() {
		WritableFont titleFont = new WritableFont(WritableFont.createFont("宋体"), 12, WritableFont.BOLD);
		WritableCellFormat titleCellFormat = new WritableCellFormat(titleFont);
		try {
			//设置标题样式：加边框、背景颜色为淡灰、居中样式
			titleCellFormat.setBorder(Border.ALL, BorderLineStyle.THIN);
			titleCellFormat.setBackground(Colour.PALE_BLUE);
			titleCellFormat.setAlignment(Alignment.CENTRE);
			titleCellFormat.setVerticalAlignment(VerticalAlignment.CENTRE);
		} catch (Exception e) {
			e.printStackTrace();
		}
		return titleCellFormat;
	}

	//设置excel头样式
	public static WritableCellFormat setHeaderCellFormat() {
		WritableFont headerFont = new WritableFont(WritableFont.createFont("宋体"), 9, WritableFont.BOLD);
		WritableCellFormat headerCellFormat = new WritableCellFormat(headerFont);
		try {
			//设置标题样式：加边框、背景颜色为淡灰、居中样式
			//设置表头样式：加边框、背景颜色为淡灰、居中样式
			headerCellFormat.setBorder(Border.ALL, BorderLineStyle.THIN);
			headerCellFormat.setBackground(Colour.PALE_BLUE);
			headerCellFormat.setAlignment(Alignment.CENTRE);
			headerCellFormat.setVerticalAlignment(VerticalAlignment.CENTRE);
		} catch (Exception e) {

		}
		return headerCellFormat;
	}

	//设置excel单元格样式
	public static WritableCellFormat setBodyCellFormat() {
		WritableFont bodyFont = new WritableFont(WritableFont.createFont("宋体"), 9, WritableFont.NO_BOLD);
		WritableCellFormat bodyCellFormat = new WritableCellFormat(bodyFont);
		try {
			//设置表格体样式：加边框、居中
			bodyCellFormat.setBorder(Border.ALL, BorderLineStyle.THIN);
			bodyCellFormat.setAlignment(Alignment.CENTRE);
			bodyCellFormat.setVerticalAlignment(VerticalAlignment.CENTRE);
		} catch (Exception e) {

		}
		return bodyCellFormat;
	}

	//设置excel单元格红色样式
	public static WritableCellFormat setBodyRedCellFormat() {
		WritableFont bodyRedFont = new WritableFont(WritableFont.createFont("宋体"), 9, WritableFont.NO_BOLD, false, UnderlineStyle.NO_UNDERLINE, Colour.RED);
		WritableCellFormat bodyRedCellFormat = new WritableCellFormat(bodyRedFont);
		try {
			//设置表格体样式：加边框、居中
			bodyRedCellFormat.setBorder(Border.ALL, BorderLineStyle.THIN);
			bodyRedCellFormat.setAlignment(Alignment.CENTRE);
			bodyRedCellFormat.setVerticalAlignment(VerticalAlignment.CENTRE);
		} catch (Exception e) {

		}
		return bodyRedCellFormat;
	}
}

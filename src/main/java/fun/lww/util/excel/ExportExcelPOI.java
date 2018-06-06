package fun.lww.util.excel;

import java.io.IOException;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.math.BigDecimal;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Collection;
import java.util.Date;
import java.util.Iterator;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;

/**
 * 
 * 利用开源组件POI3.0.2动态导出EXCEL文档
 * 
 * 转载时请保留以下信息，注明出处！
 * 
 * @author 鑫缘
 * 
 * @version v1.3
 * 
 * @param <T>
 *            应用泛型，代表任意一个符合javabean风格的类
 * 
 *            注意这里为了简单起见，boolean型的属性xxx的get器方式为getXxx(),而不是isXxx()
 * 
 *            byte[]表jpg格式的图片数据
 */

public class ExportExcelPOI<T> {

	public void exportExcel(Collection<T> dataset, OutputStream out) throws IOException, SecurityException, NoSuchMethodException {

		exportExcel("系统导出文档", null, null, dataset, out, "yyyy-MM-dd");

	}

	public void exportExcel(String[] headers, Collection<T> dataset,
			OutputStream out) throws IOException, SecurityException, NoSuchMethodException {

		exportExcel("系统导出文档", headers, null, dataset, out, "yyyy-MM-dd");

	}

	public void exportExcel(String[] headers, Collection<T> dataset,

	OutputStream out, String pattern) throws IOException, SecurityException, NoSuchMethodException {

		exportExcel("系统导出文档", headers, null, dataset, out, pattern);

	}

	public void exportExcel(String[] headers, Short[] widths,
			Collection<T> dataset, OutputStream out)  throws IOException, SecurityException, NoSuchMethodException {

		exportExcel("系统导出文档", headers, widths, dataset, out, "yy-MM-dd");

	}

	public void exportExcel(String title, String[] headers, Short[] widths,
			Collection<T> dataset, OutputStream out) throws IOException, SecurityException, NoSuchMethodException {

		exportExcel(title, headers, widths, dataset, out, "yyyy-MM-dd");

	}
	
	/**
	 * 
	 * 这是一个通用的方法，利用了JAVA的反射机制，可以将放置在JAVA集合中并且符号一定条件的数据以EXCEL 的形式输出到指定IO设备上
	 * 
	 * 
	 * 
	 * @param title
	 * 
	 *            表格标题名
	 * 
	 * @param headers
	 * 
	 *            表格属性列名数组
	 * @param widths
	 * 
	 *            表格宽度数组 默认15
	 * @param dataset
	 * 
	 *            需要显示的数据集合,集合中一定要放置符合javabean风格的类的对象。此方法支持的
	 * 
	 *            javabean属性的数据类型有基本数据类型及String,Date,byte[](图片数据)
	 * 
	 * @param out
	 * 
	 *            与输出设备关联的流对象，可以将EXCEL文档导出到本地文件或者网络中
	 * 
	 * @param pattern
	 * 
	 *            如果有时间数据，设定输出格式。默认为"yyyy-MM-dd"
	 * @throws IOException 
	 * @throws NoSuchMethodException 
	 * @throws SecurityException 
	 */

	@SuppressWarnings("unchecked")
	public void exportExcel(String title, String[] headers, Short[] widths,
			Collection<T> dataset, OutputStream out, String pattern) throws IOException, SecurityException, NoSuchMethodException {

		// 声明一个工作薄

		HSSFWorkbook workbook = new HSSFWorkbook();

		// 生成一个表格

		HSSFSheet sheet = workbook.createSheet(title);
		 

		// 设置表格默认列宽度为15个字节
		// 设置默认宽度15 字节否则循环设置宽度
		sheet.setDefaultColumnWidth((short) 15);

		if (widths != null) {
			for (int i = 0; i < widths.length; i++) {
				sheet.setColumnWidth((short) i, (short) (widths[i] * 250));
			}
		}

		// ************************************标题样式设置**************************************
		// 生成标题样式

		HSSFCellStyle style = workbook.createCellStyle();
		// 设置这些样式

		//style.setFillForegroundColor(HSSFColor.LIME.index);
		//style.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
		//
		style.setFillForegroundColor(HSSFColor.GREY_25_PERCENT.index);
		style.setFillPattern(CellStyle.SOLID_FOREGROUND);
		
		
		style.setBorderBottom(CellStyle.BORDER_THIN);

		style.setBorderLeft(CellStyle.BORDER_THIN);

		style.setBorderRight(CellStyle.BORDER_THIN);

		style.setBorderTop(CellStyle.BORDER_THIN);

		style.setAlignment(CellStyle.ALIGN_CENTER);// 左右居中

		style.setVerticalAlignment(CellStyle.VERTICAL_CENTER);// 上下居中

		style.setWrapText(true); // 自动换行

		// 生成标题字体

		HSSFFont font = workbook.createFont();

		font.setColor(HSSFColor.BLACK.index);// 字体颜色

		font.setFontHeightInPoints((short) 11);// 字号

		font.setBoldweight(Font.BOLDWEIGHT_BOLD);// 是否加粗

		// 把字体应用到当前的标题样式

		style.setFont(font);
		// ************************************数据样式设置**************************************
		// 生成并设置数据样式

		HSSFCellStyle style2 = workbook.createCellStyle();

		style2.setFillForegroundColor(HSSFColor.WHITE.index);

		style2.setFillPattern(CellStyle.SOLID_FOREGROUND);

		style2.setBorderBottom(CellStyle.BORDER_THIN);

		style2.setBorderLeft(CellStyle.BORDER_THIN);

		style2.setBorderRight(CellStyle.BORDER_THIN);

		style2.setBorderTop(CellStyle.BORDER_THIN);

		style2.setAlignment(CellStyle.ALIGN_CENTER);

		style2.setVerticalAlignment(CellStyle.VERTICAL_CENTER);

		style2.setWrapText(true); // 自动换行

		// 生成数据字体

		HSSFFont font2 = workbook.createFont();

		font2.setFontHeightInPoints((short) 9);

		font2.setBoldweight(Font.BOLDWEIGHT_NORMAL);

		// 把字体应用到当前的样式

		style2.setFont(font2);
		// ************************************数据样式设置结束**************************************
		// 声明一个画图的顶级管理器

		// HSSFPatriarch patriarch = sheet.createDrawingPatriarch();

		// 定义注释的大小和位置,详见文档

		// HSSFComment comment = patriarch.createComment(new HSSFClientAnchor(0,
		// 0, 0, 0, (short) 4, 2, (short) 6, 5));

		// 设置注释内容

		// comment.setString(new HSSFRichTextString("可以在POI中添加注释！"));

		// 设置注释作者，当鼠标移动到单元格上是可以在状态栏中看到该内容.

		// comment.setAuthor("leno");

		// 产生表格标题行
		HSSFRow row = sheet.createRow(0);
		 
		for (short i = 0; i < headers.length; i++) {

			HSSFCell cell = row.createCell(i);
//			cell.setEncoding(HSSFCell.ENCODING_UTF_16);

			cell.setCellStyle(style);

			HSSFRichTextString text = new HSSFRichTextString(headers[i]);

			cell.setCellValue(text);

		}
		
		//设置单元格公式    ：myCell.setCellFormula("SUM(A8:B8)");

		// 遍历集合数据，产生数据行
		Class tCls = null;
		List<Method> methodList = new ArrayList<Method>();
		
		if(dataset.size()>0){
			tCls = dataset.iterator().next().getClass();
			
			Field[] fields = tCls.getDeclaredFields();
			for(int i=0;i<fields.length;i++){
				
				Field field = fields[i];
				
				String fieldName = field.getName();

				String getMethodName = "get"+fieldName.substring(0, 1).toUpperCase() + fieldName.substring(1);
				
				methodList.add(tCls.getMethod(getMethodName,new Class[] {}));				
			}
			
			Iterator<T> it = dataset.iterator();	

			int index = 0;

			while (it.hasNext()) {
				index++;
				// 获取列
				row = sheet.createRow(index);
				//取出当前对象
				T t = it.next();
				//循环插入列
				for (short i = 0; i < methodList.size(); i++) {
					// 获取行
					HSSFCell cell = row.createCell(i);
					// 载入样式
					cell.setCellStyle(style2);

					try {

						//通过之前获取的方法名取出当前行对象中的当前列数据
						Object value = methodList.get(i).invoke(t, new Object[] {});

						// 判断值的类型后进行强制类型转换

						String textValue = null;


						if (value==null) {
							
							textValue = "";
							
						} else if (value instanceof String) {
							
							textValue = value.toString();
							
						} else if (value instanceof Boolean) {
							
							textValue=(Boolean) value?"是":"否";
							
						} else if (value instanceof Date) {

							Date date = (Date) value;

							SimpleDateFormat sdf = new SimpleDateFormat(pattern);

							textValue = sdf.format(date);

						} else if (value instanceof Float) {//对数值值型数据进行格式化
							DecimalFormat df = new DecimalFormat();
							df.applyPattern("#.##");
							textValue = df.format(value);
						} else if (value instanceof Integer) {
							DecimalFormat df = new DecimalFormat();
							df.applyPattern("#");
							textValue = df.format(value);
						} else if (value instanceof Double) {
							DecimalFormat df = new DecimalFormat();
							df.applyPattern("#.##");
							textValue = df.format(value);
						}else if (value instanceof BigDecimal) {
							DecimalFormat df = new DecimalFormat();
							df.applyPattern("#.##");
							textValue = df.format(value);
						}else {							
							textValue = "";							
						}

						// 如果不是图片数据，就利用正则表达式判断textValue是否全部由数字组成

						if (textValue != null) {

							Pattern p = Pattern.compile("^\\d+(\\.\\d+)?$");

							Matcher matcher = p.matcher(textValue);

							if (matcher.matches()) {

								// 是数字当作double处理

								cell.setCellValue(Double.parseDouble(textValue));

							} else {

								HSSFRichTextString richString = new HSSFRichTextString(textValue);
								
								cell.setCellValue(richString);

							}
						}

					} catch (SecurityException e) {

						// TODO Auto-generated catch block

						e.printStackTrace();

					} catch (IllegalArgumentException e) {

						// TODO Auto-generated catch block

						e.printStackTrace();

					} catch (IllegalAccessException e) {

						// TODO Auto-generated catch block

						e.printStackTrace();

					} catch (InvocationTargetException e) {

						// TODO Auto-generated catch block

						e.printStackTrace();

					} finally {
						
						// 清理资源
					}

				}

			}
			
		}
		workbook.write(out);
	}
	
}

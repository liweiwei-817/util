package fun.lww.util.excel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ImportExcelPOI {

	/**
	 * 读取 office 2003 fun.lww.util.excel
	 * @throws IOException
	 * @throws FileNotFoundException
	 */
	public String[][] read2003Excel(File file)
			throws IOException {
		List<String[]> result = new ArrayList<String[]>();
		int rowSize = 0;
		//创建新的 fun.lww.util.excel 工作博
		HSSFWorkbook hwb = new HSSFWorkbook(new FileInputStream(file));
		//在 excel工作博中建一工作表, 其名为缺省值
		// 如要新建一名为"效益指标"的工作表,其语句为:
		//HSSFSheet sheet = hwb.createSheet("");
//		HSSFSheet sheet = hwb.getSheetAt(0);
	//	System.out.println("===============    此  "+file.getFileName()+"  文档一共有  "+hwb.getNumberOfSheets()+"  列..."+hwb.getSheetAt(0).getLastRowNum()+" 行... " );
//		HSSFSheet sheet = hwb.createSheet();
		HSSFCell cell = null;
//		for (int i = sheet.getFirstRowNum(); i <= sheet.getPhysicalNumberOfRows(); i++) {
		//循环列
		for(int i=0;i<hwb.getNumberOfSheets();i++){
			HSSFSheet sheet = hwb.getSheetAt(i);
			
			//循环每行        不包括标题  导航. 如何想要把列名也显示出来， 请把下边的1  修改为0
			for(int rowIndex=1;rowIndex<=sheet.getLastRowNum();rowIndex++){
				HSSFRow row = sheet.getRow(rowIndex);//得到每一行的数据 
				if(row==null){
					continue;
				}
				//row.getLastCellNum()   是获取最后一个不为空的列是第几个
				int tempRowSize = row.getLastCellNum();
				if(tempRowSize>rowSize){
					rowSize = tempRowSize;
				}
				String[] values = new String[rowSize];
				Arrays.fill(values, "");
				boolean hasValue = false;
				for (short columnIndex = 0; columnIndex <= row.getLastCellNum(); columnIndex++) {
					String value = "";
					cell = row.getCell(columnIndex);
					if (cell == null) {
						continue;
					}
					// 注意：一定要设成这个，否则    有可能   会出现乱码
					//	cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					switch (cell.getCellType()) {
					case Cell.CELL_TYPE_STRING:
						value = cell.getStringCellValue();
						break;
					case Cell.CELL_TYPE_NUMERIC:
						/*if(HSSFDateUtil.isCellDateFormatted(cell)){
							Date date = cell.getDateCellValue();
							if(date!=null){
								value = new SimpleDateFormat("yyyy-MM-dd").format(date);
							}else{
								value = "";
							}
						}else{
							value = new DecimalFormat("##.####").format(cell.getNumericCellValue());
						}*/
						
						value = new DecimalFormat("##.####").format(cell.getNumericCellValue());
						break;
					case Cell.CELL_TYPE_FORMULA:
						// 导入时如果为公式生成的数据则无值
						if(!cell.getStringCellValue().equals("")){
							value = cell.getStringCellValue();
						}else{
							value = cell.getNumericCellValue()+"";
						}
						break;
					case Cell.CELL_TYPE_BOOLEAN:
						value = (cell.getBooleanCellValue()==true?"Y":"N");
						break;
					case Cell.CELL_TYPE_BLANK:
						value = "";
						break;
					case Cell.CELL_TYPE_ERROR:
						value = "";
						break;
					default:
						value = cell.toString();
					}
					if (columnIndex==0 && "".equals(value.trim())) {
						break;
					}
					values[columnIndex] = value.trim();
					hasValue = true;
				}
				if(hasValue){
					result.add(values);
				}
			}
		}
		String[][] returnArray = new String[result.size()][rowSize];
		for(int m=0;m<returnArray.length;m++){
			returnArray[m] = result.get(m);
		}
		return returnArray;
	}

	/**
	 * 读取Office 2007 fun.lww.util.excel
	 */
	public String[][] read2007Excel(File file) throws IOException {
		List<String[]> result = new ArrayList<String[]>();
		int rowSize = 0;
		XSSFCell cell = null;
		// 构造 XSSFWorkbook 对象，strPath 传入文件路径
		XSSFWorkbook xwb = new XSSFWorkbook(new FileInputStream(file));
		for(int sheetIndex = 0;sheetIndex<xwb.getNumberOfSheets();sheetIndex++){
			XSSFSheet sheet = xwb.getSheetAt(sheetIndex);
			for(int rowIndex=1;rowIndex<=sheet.getLastRowNum();rowIndex++){
				XSSFRow row = sheet.getRow(rowIndex);
				if(row==null){
					continue;
				}
				int tempRowSize = row.getLastCellNum();
				if (tempRowSize > rowSize) {
					rowSize = tempRowSize;
				}
				String[] values = new String[rowSize];
				Arrays.fill(values, "");
				boolean hasValue = false;
				for(short columnIndex=0;columnIndex<row.getLastCellNum();columnIndex++){
					String value = "";
					cell = row.getCell(columnIndex);//得到的每行的每个具体的值 
					if(cell!=null){
						switch (cell.getCellType()) {
						case Cell.CELL_TYPE_STRING:
							value = cell.getStringCellValue();
							break;
						case Cell.CELL_TYPE_NUMERIC:
							SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");// 格式化日期字符串
							//Double aa=cell.getNumericCellValue();
							value=new DecimalFormat("##.####").format(cell.getNumericCellValue());
							break;
						case Cell.CELL_TYPE_FORMULA:
							// 导入时如果为公式生成的数据则无值
							if (!cell.getStringCellValue().equals("")) {
								value = cell.getStringCellValue();
							} else {
								value = cell.getNumericCellValue() + "";
							}
							break;
						case Cell.CELL_TYPE_BLANK:
							break;
						case Cell.CELL_TYPE_ERROR:
							value = "";
							break;
						case Cell.CELL_TYPE_BOOLEAN:
							value = (cell.getBooleanCellValue() == true ? "Y" : "N");
							break;
						default:
							value = "";
							break;
						}
					}
					if(columnIndex==0 && value.trim().equals("")){
						break;
					}
					values[columnIndex] = value.trim();
					hasValue = true;
				}
				if(hasValue){
					result.add(values);
				}
			}
		}
		String[][] returnArray = new String[result.size()][rowSize];
		for (int i = 0; i < returnArray.length; i++) {
			returnArray[i] = result.get(i);
		}
		return returnArray;
	}
}

package fun.lww.util.excel;

import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;

import javax.servlet.http.HttpServletResponse;

import org.apache.poi.ss.formula.functions.T;

public class ExportExcelPOITest {

	public String test(HttpServletResponse response) throws IOException, SecurityException, NoSuchMethodException{
		List<T> list = new ArrayList<T>();
		String fname = "truckTipOrder";//Excel文件名
		OutputStream os = response.getOutputStream();
		//取得输出流
		response.reset();//清空输出流
		response.setHeader("Content-disposition", "attachment; filename=" + fname + ".xls");//设定输出文件头
		response.setContentType("application/msexcel");//定义输出类型
		
		ExportExcelPOI<T> ee=new ExportExcelPOI<T>();
		String[] headers={"车牌号","分单号","创建时间","创建人","分拨地点","分单状态","到达地点","代理商名称"};
		Short[] widths={10,10,10,10,10,10,10,10};
		ee.exportExcel(fname, headers, widths, list, os);
		return null;
	}
}

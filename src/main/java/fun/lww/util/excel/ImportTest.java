package fun.lww.util.excel;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

public class ImportTest {

	public void test() throws IOException {
		File file = new File("");
		String fileName = file.getName();
		List<String> list = new ArrayList<String>();
		//根据其名称获取后缀
		String extension = fileName.lastIndexOf(".") == -1 ? "" : fileName
				.substring(fileName.lastIndexOf(".") + 1);
		String[][] result = null;
		if ("xls".equals(extension)) {
			result = new ImportExcel().read2003Excel(file);
		} else if ("xlsx".equals(extension)) {
			result = new ImportExcel().read2007Excel(file);
		} else {
			throw new IOException("不支持的文件类型");
		}
	}
}

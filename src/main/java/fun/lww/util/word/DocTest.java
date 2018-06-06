package fun.lww.util.word;

import org.junit.Test;

import java.util.HashMap;
import java.util.Map;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

public class DocTest {

	@Test
	public void test1(HttpServletRequest request, HttpServletResponse response) throws Exception {
		//map中存放数据
		Map<String, String> map = new HashMap<String, String>();
		ReadAndWriteDoc.readWriteWord(response, "/word/123.doc", map , "123.doc");
	}
}

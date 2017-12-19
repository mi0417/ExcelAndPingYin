package com.codyy.pinyin;

import net.sourceforge.pinyin4j.PinyinHelper;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Scanner;

/**
 * com.codyy.pinyin
 *
 * @author suyue
 * @version 1.0
 * @Description: 获得学校域名和用户名
 * @date 2017/12/19 9:59
 */
public class GetSchools {

	/**
	 * 得到中文首字母
	 *
	 * @param str
	 *
	 * @return
	 */
	public static String getPinYinHeadChar(String str) {

		String convert = "";
		for (int j = 0; j < str.length(); j++) {
			char word = str.charAt(j);
			String[] pinyinArray = PinyinHelper.toHanyuPinyinStringArray(word);
			if (pinyinArray != null) {
				convert += pinyinArray[0].charAt(0);
			} else {
				convert += word;
			}
		}
		return convert;
	}


	public Boolean getDomainAndUserName(String xmlPath, int a, int b, int c, int d, int column) {
		if (xmlPath == null || "".equals(xmlPath)) {

			System.out.println("未传入xml名称或xml名称为空!");
			return false;
		}

		List<String> schools = new ArrayList<>();    //学校名称列表
		List<String> users = new ArrayList<>();    //用户姓名列表

		List<String> domains = new ArrayList<>();    //学校域名列表
		List<String> userNames = new ArrayList<>();    //用户名列表

		String path = System.getProperty("user.dir");
		File file = new File(path + "/" + xmlPath);

		if (!file.exists()) {
			System.out.println(path + "\\" + xmlPath + " 不存在!");
			try {
				Thread.sleep(3000L);
			} catch (InterruptedException e) {
				e.printStackTrace();
			}
			return false;
		}

		FileInputStream in = null;
		FileOutputStream out = null;
		try {
			in = new FileInputStream(file);
			//得到Excel工作簿对象
			HSSFWorkbook wb = new HSSFWorkbook(in);
			//得到Excel工作表对象
			HSSFSheet sheet = wb.getSheetAt(0);
			int firstRowNum = sheet.getFirstRowNum();
			int lastRowNum = sheet.getLastRowNum();

			Row row = null;
			Cell cell_a = null;
			Cell cell_b = null;
			Cell cell_c = null;
			Cell cell_d = null;
			for (int i = firstRowNum + column; i <= lastRowNum; i++) {
				row = sheet.getRow(i);          //取得第i行
				cell_a = row.getCell(a);        //取得i行的第一列
				String cellValue_a = cell_a.getStringCellValue().trim();
				schools.add(cellValue_a);

				cell_b = row.getCell(b);
				String cellValue_b = cell_b.getStringCellValue().trim();
				users.add(cellValue_b);

				cell_c = row.getCell(c);    //域名
				cell_d = row.getCell(d);    //用户名

				String domain = GetSchools.getPinYinHeadChar(cellValue_a);
				while (true) {
					Integer n = 1;
					if (domains.contains(domain)) {
						domain = domain + n;
						n++;
					} else {
						domains.add(domain);
						break;
					}
				}
				if (cell_c == null) {
					row.createCell(c);
					cell_c = row.getCell(c);
				}
				cell_c.setCellValue("http://" + domain + ".needu.cn");

				String userName = GetSchools.getPinYinHeadChar(cellValue_b);
				while (true) {
					Integer n = 1;
					if (userNames.contains(userName)) {
						userName = userName + n;
						n++;
					} else {
						userNames.add(userName);
						break;
					}
				}
				if (cell_d == null) {
					row.createCell(d);
					cell_d = row.getCell(d);
				}
				cell_d.setCellValue(userName);


				System.out.println(cellValue_a + "----" + domain + "----" + cellValue_b + "----" + userName);
			}

			out = new FileOutputStream(file);
			wb.write(out);

			System.out.println("已成功生成!");

			return true;
		} catch (Exception e) {
			System.out.println(e.getMessage());
			return false;
		} finally {
			try {
				in.close();
			} catch (IOException e) {
//				e.printStackTrace();
			}
			try {
				out.close();
			} catch (IOException e) {
//				e.printStackTrace();
			}
		}
	}

	public static void main(String[] args) {
		Scanner input = new Scanner(System.in);

		System.out.println("注意事项:\n" +
				"1、默认学校名在第A列，用户姓名在第I列，域名在第D列，用户名在第H列\n" +
				"2、默认数据从第3行开始\n" +
				"3、Excel放在当前目录下,名称应为schools.xls\n" +
				"按Y开始");

		input.nextLine();

		int a = 0;    //学校名列数
		int b = 8;    //用户姓名列数
		int c = 3;    //域名列数
		int d = 7;    //用户名列数
		int column = 2;	//数据前有两行

		GetSchools getSchools = new GetSchools();

		String xmlPath = "schools.xls";

		getSchools.getDomainAndUserName(xmlPath, a, b, c, d, column);


	}
}

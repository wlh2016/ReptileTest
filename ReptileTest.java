package com.reptile.test;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;

public class ReptileTest {
	
	private int serialNum = 1;

	public void reptileTestForBooks(String bookName, int n) {
		try {
			//对每一页的数据进行抓取
			Document doc = Jsoup.connect("https://book.douban.com/tag/" + bookName + "?start=" + n + "&type=S").get();
			Elements books = doc.getElementsByClass("subject-item");
			String assess;
			int assessCount;
			String[] bookInfo = new String[5];
			//获取excel文件并对其进行写入
			HSSFWorkbook hwb = new HSSFWorkbook(new FileInputStream("豆瓣书籍统计.xls"));
			HSSFSheet sheet = hwb.getSheetAt(0);
			sheet.autoSizeColumn(1);
			sheet.autoSizeColumn(4);
			sheet.autoSizeColumn(5);
			for(Element book:books) {
				// 获取评价数，若不低于1000则保存到excel中
				assess = book.getElementsByClass("pl").html();
				if(!assess.startsWith("(目前") && !assess.startsWith("(少于")) {
					assessCount = Integer.parseInt(assess.substring(1, assess.length()-4));
					if(assessCount >= 1000) {
						HSSFRow row = sheet.getRow(this.serialNum);
						row = sheet.createRow((short) sheet.getLastRowNum()+1);
						//序号
						row.createCell(0).setCellValue(this.serialNum);
						//书名
						row.createCell(1).setCellValue(book.getElementsByTag("a").attr("title"));
						//评分
						row.createCell(2).setCellValue(book.getElementsByClass("rating_nums").html());
						bookInfo = book.getElementsByClass("pub").html().split("/");
						//评价人数
						row.createCell(3).setCellValue(assessCount);
						if(bookInfo.length == 5) {
							//作者
							row.createCell(4).setCellValue(bookInfo[0]);
							//出版社
							row.createCell(5).setCellValue(bookInfo[2]);
							//出版日期
							row.createCell(6).setCellValue(bookInfo[3]);
							//价格
							row.createCell(7).setCellValue(bookInfo[4]);
						} else {
							//作者
							row.createCell(4).setCellValue(bookInfo[0]);
							//出版社
							row.createCell(5).setCellValue(bookInfo[1]);
							//出版日期
							row.createCell(6).setCellValue(bookInfo[2]);
							//价格
							row.createCell(7).setCellValue(bookInfo[3]);
						}
						this.serialNum++;
					}
				}
			}
			FileOutputStream out = null;
			try {
				out = new FileOutputStream("豆瓣书籍统计.xls");
				hwb.write(out);
			} catch (FileNotFoundException e) {
				e.printStackTrace();
			} finally {
				out.close();
			}
			if(n <= 1000) {
				reptileTestForBooks(bookName, n + 20);
			} else {
				return;
			}
		} catch (IOException e) {
			e.printStackTrace();
		}
	}
	
	public static void main(String[] args) {
		//新建一个excel
		HSSFWorkbook hwb = new HSSFWorkbook();
		HSSFSheet sheet = hwb.createSheet("豆瓣书籍统计");
		HSSFRow firstRow = sheet.createRow(0);
		HSSFCell[] firstCell = new HSSFCell[8];
		String[] names = new String[]{"序号", "书名", "评分", "评价人数", "作者", "出版社", "出版日期", "价格"};
		for(int i=0, length=names.length; i<length; i++) {
			firstCell[i] = firstRow.createCell(i);
			firstCell[i].setCellValue(new HSSFRichTextString(names[i]));
		}
		try {
			OutputStream os = new FileOutputStream("豆瓣书籍统计.xls");
			hwb.write(os);
			os.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
		ReptileTest reptileTest = new ReptileTest();
		reptileTest.reptileTestForBooks("互联网", 0);
		reptileTest.reptileTestForBooks("编程", 0);
		reptileTest.reptileTestForBooks("算法", 0);
		
	}

}

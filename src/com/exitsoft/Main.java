package com.exitsoft;

import jxl.Workbook;
import jxl.format.Alignment;
import jxl.format.Border;
import jxl.format.BorderLineStyle;
import jxl.write.Label;
import jxl.write.WritableCellFormat;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;

import java.io.File;
import java.util.*;

public class Main {

    final static String OPERA_USER_AGENT = "Opera/9.80 (Windows NT 6.1; U; ko) Presto/2.6.30 Version/10.62";
    public ArrayList<String> title_list = new ArrayList<String>();
    public static String query;
    public static String ds;
    public static String de;

    public static List<Map<String, String>> list = new ArrayList<Map<String, String>>();

    public static void main(String[] args) {
        Scanner s = new Scanner(System.in);
        System.out.println("검색어를 입력해주세요");
        query = s.next();
        System.out.println("시작날짜를 입력해주세요");
        ds = s.next();
        System.out.println("종료날짜를 입력해주세요");
        de = s.next();
        System.out.println("저장하고 싶은 이름을 입력하세요");
        String save_name = s.next();

        String url = "https://search.naver.com/search.naver?ie=utf8&where=news&query=";
        String data = "&sm=tab_pge&sort=2&photo=0&field=0&reporter_article=&pd=3&ds="+ds+"&de="+de+"&docid=&mynews=0&start=";


        try {
            for(int i = 0; i < 1; i++){
                String real_url = url + query + data;
//                real_url = "https://search.naver.com/search.naver?ie=utf8&where=news&query=%ED%95%9C%EC%96%91%EB%8C%80&sm=tab_pge&sort=2&photo=0&field=0&reporter_article=&pd=3&ds=2015.01.01&de=2016.01.01&docid=&mynews=0&start=11&refresh_start=0";
                System.out.println(real_url+ String.valueOf(i*10+1) + "&refresh_start=0");
                real_url = real_url + String.valueOf(i*10+1) + "&refresh_start=0";
                Document doc = Jsoup.connect(real_url)
                        .userAgent(OPERA_USER_AGENT)
                        .header("Accept-Language", "ko-KR,ko;q=0.8,en-US;q=0.6,en;q=0.4")
                        .get();
                Elements titles = doc.select("a._sp_each_title");
                Elements news_agent = doc.select("span._sp_each_source");
                Elements date = doc.select("dd.txt_inline");
                for(int j = 0; j < titles.size(); j++){
                    HashMap<String, String> map = new HashMap<String, String>();
                    map.put("title", titles.get(j).text());
                    map.put("link", titles.get(j).attr("href"));
                    map.put("news_agent", news_agent.get(j).text());
                    map.put("date", date.get(j).text().split(" ")[1]);

                    list.add(map);

                    System.out.println(titles.get(j).text());
                    System.out.println(titles.get(j).attr("href"));
                    System.out.println(news_agent.get(j).text());
                    System.out.println(date.get(j).text().split(" ")[1]);
                }
            }

            excel_output(save_name);


        }
        catch (Exception ex){
            System.out.println(ex.toString());
        }
	// write your code here

    }


    public static void excel_output(String save_name){
        try {
            WritableWorkbook wb = Workbook.createWorkbook(new File(save_name+".xls"));
            // WorkSheet 생성
            WritableSheet sh = wb.createSheet("네이버", 0);

            // 열넓이 설정 (열 위치, 넓이)
            sh.setColumnView(0, 20);
            sh.setColumnView(1, 20);
            sh.setColumnView(2, 15);
            sh.setColumnView(3, 50);

            // 셀형식
            WritableCellFormat textFormat = new WritableCellFormat();
            // 생성
            textFormat.setAlignment(Alignment.LEFT);
            // 테두리
            textFormat.setBorder(Border.ALL, BorderLineStyle.THIN);

            int row = 0;

            Label label = new jxl.write.Label(0, row, "검색어", textFormat);
            sh.addCell(label);

            label = new jxl.write.Label(1, row, query, textFormat);
            sh.addCell(label);

            label = new jxl.write.Label(2, row, "시작날짜", textFormat);
            sh.addCell(label);

            label = new jxl.write.Label(3, row, ds, textFormat);
            sh.addCell(label);

            label = new jxl.write.Label(4, row, "종료날짜", textFormat);
            sh.addCell(label);

            label = new jxl.write.Label(5, row, de, textFormat);
            sh.addCell(label);

            row++;

            // 헤더
            label = new jxl.write.Label(0, row, "제목", textFormat);
            sh.addCell(label);

            label = new jxl.write.Label(1, row, "링크", textFormat);
            sh.addCell(label);

            label = new jxl.write.Label(2, row, "날짜", textFormat);
            sh.addCell(label);

            label = new jxl.write.Label(3, row, "뉴스사", textFormat);
            sh.addCell(label);

            row++;

            for (Map<String, String> tem : list) {

                // 이름
                label = new jxl.write.Label(0, row, tem.get("title"),
                        textFormat);
                sh.addCell(label);

                // 링크
                label = new jxl.write.Label(1, row, tem.get("link"),
                        textFormat);
                sh.addCell(label);

                // 날짜
                label = new jxl.write.Label(2, row, tem.get("date"),
                        textFormat);
                sh.addCell(label);


                // 뉴스사
                label = new jxl.write.Label(3, row, tem.get("news_agent"),
                        textFormat);
                sh.addCell(label);

                row++;
            }
            // WorkSheet 쓰기
            wb.write();

            // WorkSheet 닫기
            wb.close();
        }
        catch (Exception ex){
            System.out.println(ex);
        }


    }
}

package com.exitsoft;

import jxl.Workbook;
import jxl.format.Alignment;
import jxl.format.Border;
import jxl.format.BorderLineStyle;
import jxl.write.*;

import javax.swing.*;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.util.Map;

/**
 * Created by nayunhwan on 2016. 2. 10..
 */

public class Frame extends JFrame{
    JPanel pan = new JPanel(new GridLayout(6,1));
    JPanel pan_btn = new JPanel(new GridLayout(1,2));

    JLabel label_query = new JLabel("검색어를 입력해주세요");
    JLabel label_ds = new JLabel("시작 연도를 입력해주세요\n(ex 2016.01.01)");
    JLabel label_de = new JLabel("종료 연도를 입력해주세요\n(ex 2016.01.01)");

    JTextField input_query = new JTextField();
    JTextField input_ds = new JTextField();
    JTextField input_de = new JTextField();


    JButton btn_ok = new JButton("확인");
    JButton btn_cancel = new JButton("취소");



    public Frame(){
        setTitle("뉴스 기사 검색");
        setDefaultCloseOperation(DISPOSE_ON_CLOSE);
        setLayout(new BorderLayout());

        pan.add(label_query);
        pan.add(input_query);
        pan.add(label_ds);
        pan.add(input_ds);
        pan.add(label_de);
        pan.add(input_de);

        pan_btn.add(btn_ok);
        pan_btn.add(btn_cancel);

        add(pan);
        add(pan_btn, "South");



        btn_ok.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                JFrame parentFrame = new JFrame();
                JFileChooser fileChooser = new JFileChooser();
                fileChooser.setDialogTitle("Save file");
                fileChooser.addChoosableFileFilter(new MyFilter(".xlsx", "xlsx"));
                int userSelection = fileChooser.showSaveDialog(null);

                // 확인버튼을 눌렀을 때
                if(userSelection == JFileChooser.APPROVE_OPTION){
                    System.out.println("저장 경로 : " + fileChooser.getSelectedFile().toString() + "." + fileChooser.getFileFilter().getDescription());
                }
            }
        });
        setSize(300,300);
        setVisible(true);

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

            jxl.write.Label label = new jxl.write.Label(0, row, "검색어", textFormat);
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

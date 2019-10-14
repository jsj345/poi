package com.baizhi;

import com.baizhi.entity.User;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.junit.Test;
import org.junit.runner.RunWith;
import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.test.context.junit4.SpringRunner;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

@RunWith(SpringRunner.class)
@SpringBootTest
public class PoiApplicationTests {

    @Test
    public void contextLoads() {
        HSSFWorkbook sheets = new HSSFWorkbook();
        HSSFSheet sheet = sheets.createSheet("测试");
        HSSFRow row = sheet.createRow(0);
        HSSFCell cell = row.createCell(0);
        cell.setCellValue("aaaa");
        try {
            sheets.write(new FileOutputStream(new File("D:/a.xls")));
            System.out.println("创建成功");
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    @Test
    public void aaaa() {
        HSSFWorkbook sheets = new HSSFWorkbook();
        HSSFDataFormat dataFormat = sheets.createDataFormat();
        short format = dataFormat.getFormat("yyyy年mm月dd日");
        HSSFCellStyle cellStyle = sheets.createCellStyle();
        cellStyle.setDataFormat(format);
        HSSFCellStyle cellStyle1 = sheets.createCellStyle();
        cellStyle1.setAlignment(HorizontalAlignment.CENTER);
        HSSFFont font = sheets.createFont();
        font.setBold(true);
        font.setFontName("微软雅黑");
        font.setColor(Font.COLOR_RED);
        font.setItalic(true);
        cellStyle1.setFont(font);
        HSSFSheet sheet = sheets.createSheet("测试");
        sheet.setColumnWidth(2, 15 * 256);
        HSSFRow row = sheet.createRow(0);
        String[] s = {"id", "姓名", "生日"};
        for (int i = 0; i < s.length; i++) {
            HSSFCell cell = row.createCell(i);
            cell.setCellStyle(cellStyle1);
            cell.setCellValue(s[i]);
        }
        List<User> users = new ArrayList<>();
        User user = new User("1", "aaa", new Date());
        User user1 = new User("2", "bbb", new Date());
        User user2 = new User("3", "ccc", new Date());
        users.add(user);
        users.add(user1);
        users.add(user2);
        for (int i = 0; i < users.size(); i++) {
            HSSFRow row1 = sheet.createRow(i + 1);
            HSSFCell cell = row1.createCell(0);
            cell.setCellValue(users.get(i).getId());
            HSSFCell cell1 = row1.createCell(1);
            cell1.setCellValue(users.get(i).getName());
            HSSFCell cell2 = row1.createCell(2);
            cell2.setCellStyle(cellStyle);
            cell2.setCellValue(users.get(i).getBir());
            try {
                sheets.write(new FileOutputStream(new File("D:/a.xls")));
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }
}

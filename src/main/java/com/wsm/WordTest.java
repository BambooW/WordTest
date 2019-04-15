package com.wsm;


import com.lowagie.text.*;
import com.lowagie.text.Font;
import com.lowagie.text.Image;
import com.lowagie.text.pdf.BaseFont;
import com.lowagie.text.rtf.RtfWriter2;

import java.awt.*;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.net.MalformedURLException;
import java.util.ArrayList;
import java.util.List;

/**
 * @name: WordTest
 * @Author: wangshimin
 * @Date: 2019/4/15  10:46
 * @Description:
 */
public class WordTest {
    public static void main(String[] args) {
        exportDoc();
    }

    /**
     *
     * @Description: 将网页内容导出为word @param @param file @param @throws
     *               DocumentException @param @throws IOException 设定文件 @return void
     *               返回类型 @throws
     */
    public static String exportDoc() {
        try {
            // 设置纸张大小
            Document document = new Document(PageSize.A4);
            // 建立一个书写器(Writer)与document对象关联，通过书写器(Writer)可以将文档写入到磁盘中
            // ByteArrayOutputStream baos = new ByteArrayOutputStream();
            // C:\\Users\\orion\\Desktop\\home.jpg
            File file = new File("f://2003-1.doc");
            RtfWriter2.getInstance(document, new FileOutputStream(file));
            document.open();
            // 设置中文字体
            BaseFont bfChinese = BaseFont.createFont(BaseFont.HELVETICA, BaseFont.WINANSI, BaseFont.NOT_EMBEDDED);
            // 标题字体风格
            // Font titleFont = new Font(bfChinese, 12, Font.BOLD);
            // Paragraph title = new Paragraph("测试结果");
            // 设置标题格式对齐方式
            // title.setAlignment(Element.ALIGN_CENTER);
            // title.setFont(titleFont);
            // document.add(title);
            // 正文字体风格
            Font contextFont = new Font(bfChinese, 12, Font.BOLD);
            // Font contextFont = new Font(bfChinese, 11, Font.NORMAL);
            List<String> list = new ArrayList<>();
            list.add("a");
            list.add("a");
            list.add("a");
            list.add("a");
            list.add("a");
            list.add("a");
            list.add("a");
            list.add("a");
            list.add("a");
            for (String string : list) {
                // code
                String code = "code ：  ";
                Paragraph codeStyle = new Paragraph(code);
                // 正文格式左对齐
                codeStyle.setAlignment(Element.ALIGN_LEFT);
                codeStyle.setFont(contextFont);
                // 离上一段落（标题）空的行数
                codeStyle.setSpacingBefore(20);
                // 设置第一行空的列数
                // context.setFirstLineIndent(0);
                document.add(codeStyle);
                // 操作
                String codeContent = "点击";
                Paragraph codeContentStyle = new Paragraph(codeContent, FontFactory
                        .getFont(FontFactory.HELVETICA_BOLDOBLIQUE, 11, Font.UNDERLINE, new Color(0, 0, 255)));
                // 离上一段落（标题）空的行数
                codeContentStyle.setSpacingBefore(5);
                document.add(codeContentStyle);
                // result
                String result = "result ：";
                Paragraph resultStyle = new Paragraph(result);
                // 正文格式左对齐
                resultStyle.setAlignment(Element.ALIGN_LEFT);
                resultStyle.setFont(contextFont);
                // 离上一段落空的行数
                resultStyle.setSpacingBefore(10);
                // 设置第一行空的列数
                // context.setFirstLineIndent(0);
                document.add(resultStyle);
                // 利用类FontFactory结合Font和Color可以设置各种各样字体样式
                // 结果
                String resultContent = "成功";
                Paragraph resultContentStyle = null;
                if (resultContent.equals("成功")) {
                    resultContentStyle = new Paragraph(resultContent, FontFactory
                            .getFont(FontFactory.HELVETICA_BOLDOBLIQUE, 11, Font.UNDERLINE, new Color(0, 255, 0)));
                } else {
                    resultContentStyle = new Paragraph(resultContent, FontFactory
                            .getFont(FontFactory.HELVETICA_BOLDOBLIQUE, 11, Font.UNDERLINE, new Color(255, 0, 0)));
                }
                // 离上一段落空的行数
                resultContentStyle.setSpacingBefore(5);
                document.add(resultContentStyle);
                // 添加图片 Image.getInstance即可以放路径又可以放二进制字节流
                String imgPath = "f://abc.jpg";
                Image img = Image.getInstance(imgPath);
                img.setAbsolutePosition(0, 0);
                img.setAlignment(Image.ALIGN_CENTER);// 设置图片显示位置
                img.scalePercent(30);// 表示显示的大小为原尺寸的50%
                // img.scaleAbsolute(60, 60);// 直接设定显示尺寸
                // img.scalePercent(25, 12);//图像高宽的显示比例
                // img.setRotation(30);//图像旋转一定角度
                document.add(img);
                String log = "aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa";
                if (log != null && !"".equals(log)) {
                    Paragraph exceptionStyle = new Paragraph("异常信息 ：");
                    // 正文格式左对齐
                    exceptionStyle.setAlignment(Element.ALIGN_LEFT);
                    exceptionStyle.setFont(contextFont);
                    // 离上一段落空的行数
                    exceptionStyle.setSpacingBefore(20);
                    document.add(exceptionStyle);
                    document.add(new Paragraph(log,
                            FontFactory.getFont(FontFactory.HELVETICA_BOLDOBLIQUE, 10, Font.NORMAL)));
                }
            }
            document.add(new Paragraph("生成表格"));
            /*
             * 创建有三列的表格
             */
            Table table = new Table(5);//设置表格列数
            table.setBorderWidth(1);
            table.setBorderColor(Color.BLACK);
            table.setPadding(0);
            table.setSpacing(0);

            /*
             * 添加表头的元素
             */
            Cell cell = new Cell();//单元格
            cell.setColspan(5);//设置表格为三列
            cell.setRowspan(2);//设置表格为三行

            // 表格的主体
            cell = new Cell();
            cell.setRowspan(2);//当前单元格占两行,纵向跨度
            //table.addCell(cell);
            table.addCell("1,1");
            table.addCell("1,2");
            table.addCell("1,3");
            table.addCell("1,4");
            table.addCell("1,5");
            table.addCell(new Paragraph("用java生成的表格1"));
            table.addCell(new Paragraph("用java生成的表格2"));
            table.addCell(new Paragraph("用java生成的表格3"));
            table.addCell(new Paragraph("用java生成的表格4"));
            table.addCell(new Paragraph("用java生成word文件"));
            document.add(table);
            document.close();
            // 得到输入流
            // wordFile = new ByteArrayInputStream(baos.toByteArray());
            // baos.close();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (BadElementException e) {
            e.printStackTrace();
        } catch (MalformedURLException e) {
            e.printStackTrace();
        } catch (DocumentException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
        return "";

    }
}

package com.abloz;


import com.itextpdf.text.BaseColor;
import com.itextpdf.text.DocumentException;
import org.junit.Assert;
import org.junit.Test;

import java.awt.*;
import java.io.IOException;


public class WaterMarkUtilTest {

    static String path="c:/zhh/";
    @Test
    public void setWaterMarkFromStr() {
        WaterMarkUtil.setWaterMarkFromTxt(path+"id.png",path+"id2.png", "我的水印",
                Color.WHITE,50,"黑体",100,100,"png");
    }

    @Test
    public void createWaterMarkImage() throws IOException {
        ImageUtil.createImageFromTxt("我的水印",path+"mywater.png",300,300,
                Color.GREEN,50,"黑体","png");
    }

    @Test
    public void setWaterMarkFromImg() {
        WaterMarkUtil.setWaterMarkFromImg(path+"id.png",path+"idimg2.png",
                path+"mywater.png",30,0.1f,30,30,"png");
    }

    @Test
    public void setPdfWatermark() {
        try {
            WaterMarkUtil.setPdfWatermark(path+"pdf1.pdf", path+"pdf2.pdf", "我的水印",0.5f, BaseColor.WHITE);
        } catch (
                DocumentException e) {
            e.printStackTrace();
        } catch (
                IOException e) {
            e.printStackTrace();
        }
    }

    @Test
    public void setWordWaterMark() {
        try {
            WaterMarkUtil.setWordWaterMark(path+"word1.docx", path+"word2.docx", "我是水印","#00ff00");
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    @Test
    public void putWaterRemarkToExcel() {
    }

    @Test
    public void setExcelWaterMark() {
        try {
            WaterMarkUtil.setExcelWaterMark(path+"excel1.xlsx", path+"excel2.xlsx", "我是水印","c:/zhh");
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    @Test
    public void getFileType() {
        String fileType = ImageUtil.getFileType("myfile.PNG");
        Assert.assertEquals(fileType,"png");
    }

    @Test
    public void setPPTWaterMark() {
        try {
            WaterMarkUtil.setPPTWaterMark(path+"ppt1.pptx", path+"ppt2.pptx", "我的水印","c:/zhh");

        } catch (IOException e) {
            e.printStackTrace();
        }catch (Exception e) {
            e.printStackTrace();
        }
    }
}

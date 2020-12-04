package com.abloz;
/**
 * image util for water mark
 * Author: zhouhh <ablozhou@gmail.com>
 * Date:2020/12/4
 */
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import javax.imageio.ImageIO;
import java.awt.*;
import java.awt.font.FontRenderContext;
import java.awt.geom.Rectangle2D;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.IOException;
import java.text.DecimalFormat;

public class ImageUtil {
    static Logger logger = LoggerFactory.getLogger(WaterMarkUtil.class);
    final static int mb = 1000*1000;
    public static String getFileType(String filePath) {
        int index = filePath.lastIndexOf(".");
        String fileType=filePath.substring(index+1).toLowerCase();
        logger.debug("input file:"+filePath+",type:"+fileType);
        return fileType;
    }

    /**
     * compress image
     * @param srcFile srcFile
     * @param dstFile compressed file
     * @param length
     * @param newRate
     * @param formatName
     * @throws IOException
     */
    public static void zoom(String srcFile, String dstFile,long length,double newRate, String formatName) throws IOException {
        DecimalFormat df = new DecimalFormat("0.00");// 设置保留位数
        double rate=0.8;
        logger.info("src image size：" + df.format((float) length / mb) + "MB");
        long newfile = new File(srcFile).length();
        int i = 1;
        // 如果首次压缩还大于2MB则继续处理
        while ((float) newfile / mb >= newRate) {
            logger.info("compress size：" + newfile);
            rate = rate - 0.05;// 暂定按照0.03频率压缩
            logger.info(i + " rate=" + rate);
            BufferedImage srcImage = ImageIO.read(new File(srcFile));
            int WIDTH = (int) (srcImage.getWidth() * rate);
            int HEIGHT = (int) (srcImage.getHeight() * rate);
            BufferedImage image = new BufferedImage(WIDTH, HEIGHT, BufferedImage.TYPE_INT_RGB);
            Graphics g = image.getGraphics();
            g.drawImage(srcImage, 0, 0, WIDTH, HEIGHT, null);
            // 缩小
            ImageIO.write(image, formatName, new File(dstFile));
            i++;
            newfile = new File(dstFile).length();
            logger.info("compress time：" + i);
        }
        // 调整方向
        BufferedImage newImage = ImageIO.read(new File(dstFile));
        BufferedImage image1 = Rotate(newImage, 90);// 顺时针旋转90度
        ImageIO.write(image1, formatName, new File(dstFile));
        logger.info("final path：" + dstFile + ";size："
                + df.format((float) new File(dstFile).length() / mb) + "MB");
    }

    /**
     * 对图片进行旋转
     *
     * @param src
     *            被旋转图片
     * @param angel
     *            旋转角度
     * @return 旋转后的图片
     */
    public static BufferedImage Rotate(Image src, int angel) {
        int src_width = src.getWidth(null);
        int src_height = src.getHeight(null);
        // 计算旋转后图片的尺寸
        Rectangle rect_des = CalcRotatedSize(new Rectangle(new Dimension(src_width, src_height)), angel);
        BufferedImage res = null;
        res = new BufferedImage(rect_des.width, rect_des.height, BufferedImage.TYPE_INT_RGB);
        Graphics2D g2 = res.createGraphics();
        // 进行转换
        g2.translate((rect_des.width - src_width) / 2, (rect_des.height - src_height) / 2);
        g2.rotate(Math.toRadians(angel), src_width / 2, src_height / 2);

        g2.drawImage(src, null, null);
        return res;
    }

    /**
     * 计算旋转后的图片
     *
     * @param src
     *            被旋转的图片
     * @param angel
     *            旋转角度
     * @return 旋转后的图片
     */
    public static Rectangle CalcRotatedSize(Rectangle src, int angel) {
        // 如果旋转的角度大于90度做相应的转换
        if (angel >= 90) {
            if (angel / 90 % 2 == 1) {
                int temp = src.height;
                src.height = src.width;
                src.width = temp;
            }
            angel = angel % 90;
        }

        double r = Math.sqrt(src.height * src.height + src.width * src.width) / 2;
        double len = 2 * Math.sin(Math.toRadians(angel) / 2) * r;
        double angel_alpha = (Math.PI - Math.toRadians(angel)) / 2;
        double angel_dalta_width = Math.atan((double) src.height / src.width);
        double angel_dalta_height = Math.atan((double) src.width / src.height);

        int len_dalta_width = (int) (len * Math.cos(Math.PI - angel_alpha - angel_dalta_width));
        int len_dalta_height = (int) (len * Math.cos(Math.PI - angel_alpha - angel_dalta_height));
        int des_width = src.width + len_dalta_width * 2;
        int des_height = src.height + len_dalta_height * 2;
        return new Rectangle(new Dimension(des_width, des_height));
    }
    /**
     * 根据文字生成水印图片
     * @param content
     * @param path 文件名，包含全路径
     * @param color 设置字体颜色和透明度 Color.white,new Color(255, 180, 0, 80)
     * @param width 宽
     * @param height 高
     * @param fontSize 字号大小 50
     * @param fontName 字体名称 如 宋体
     * @param formatName 输出文件格式 JPG, JPEG, PNG, GIF, BMP
     * @throws IOException
     */
    public static void createImageFromTxt(String content, String path,int width, int height, Color color,
                                            int fontSize, String fontName, String formatName) throws IOException {
        BufferedImage image = new BufferedImage(width, height, BufferedImage.TYPE_INT_RGB);// 获取bufferedImage对象

        Integer fontStyle = Font.PLAIN;

        Font font = new Font(fontName, fontStyle, fontSize);
        Graphics2D g2d = image.createGraphics(); // 获取Graphics2d对象
        image = g2d.getDeviceConfiguration().createCompatibleImage(width, height, Transparency.TRANSLUCENT);
        g2d.dispose();
        g2d = image.createGraphics();
        g2d.setColor(color);
        g2d.setStroke(new BasicStroke(1)); // 设置字体
        g2d.setFont(font); // 设置字体类型  加粗 大小
        g2d.rotate(Math.toRadians(-10), (double) image.getWidth() / 2, (double) image.getHeight() / 2);//设置倾斜度
        FontRenderContext context = g2d.getFontRenderContext();
        Rectangle2D bounds = font.getStringBounds(content, context);
        double x = (width - bounds.getWidth()) / 2;
        double y = (height - bounds.getHeight()) / 2;
        double ascent = -bounds.getY();
        double baseY = y + ascent;
        // 写入水印文字原定高度过小，所以累计写水印，增加高度
        g2d.drawString(content, (int) x, (int) baseY);
        // 设置透明度
        g2d.setComposite(AlphaComposite.getInstance(AlphaComposite.SRC_OVER));
        // 释放对象
        g2d.dispose();
        ImageIO.write(image, formatName, new File(path));
    }
}

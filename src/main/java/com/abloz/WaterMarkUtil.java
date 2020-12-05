package com.abloz;
/**
 * water mark util,support images ,pdf, powerpoint, word,excel file format to add water mark.
 * Author: zhouhh <ablozhou@gmail.com>
 * Date:2020/12/4
 * copy right 2020 zhouhh
 */

import com.itextpdf.text.BaseColor;
import org.apache.poi.util.IOUtils;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import javax.imageio.ImageIO;
import java.awt.*;

import java.awt.image.BufferedImage;
import java.io.*;

import org.apache.poi.sl.usermodel.PictureData;
import org.apache.poi.sl.usermodel.PictureData.PictureType;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Picture;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFPictureShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.model.XWPFHeaderFooterPolicy;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFHeader;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.xmlbeans.XmlObject;
import org.openxmlformats.schemas.presentationml.x2006.main.CTSlide;

import com.itextpdf.text.DocumentException;
import com.itextpdf.text.Element;
import com.itextpdf.text.pdf.BaseFont;
import com.itextpdf.text.pdf.PdfContentByte;
import com.itextpdf.text.pdf.PdfGState;
import com.itextpdf.text.pdf.PdfReader;
import com.itextpdf.text.pdf.PdfStamper;
public class WaterMarkUtil {
    static Logger logger = LoggerFactory.getLogger(WaterMarkUtil.class);

    /**
     * 添加文本水印
     * @param srcImgPath 源文件名，包含路径
     * @param dstImgPath 添加水印后输出文件，包含路径
     * @param waterMark 添加的水印文字
     * @param color 水印颜色
     * @param fontSize 字号大小
     * @param fontName 字体名称 如 宋体
     * @param offsetX 偏移x default 20
     * @param offsetY 偏移y default 20
     * @param formatName 输出图片格式 JPG, JPEG, PNG, GIF, BMP
     * @return -1 添加失败，0 添加成功
     */
    public static int setWaterMarkFromTxt(String srcImgPath, String dstImgPath, String waterMark,
                                          Color color, int fontSize, String fontName, int offsetX, int offsetY, String formatName) {
        // 原图位置, 输出图片位置, 水印文字颜色, 水印文字
        try {
            // 读取原图片信息
            File srcImgFile = new File(srcImgPath);
            Image srcImg = ImageIO.read(srcImgFile);
            int srcImgWidth = srcImg.getWidth(null);
            int srcImgHeight = srcImg.getHeight(null);
            // 加水印
            BufferedImage bufImg = new BufferedImage(srcImgWidth,
                    srcImgHeight,
                    BufferedImage.TYPE_INT_RGB);
            //获取 Graphics2D 对象
            Graphics2D g = bufImg.createGraphics();
            //设置绘图区域
            g.drawImage(srcImg, 0, 0, srcImgWidth, srcImgHeight, null);
            //设置字体
            Font font = new Font(fontName, Font.PLAIN, fontSize);

            // 根据图片的背景设置水印颜色
            g.setColor(color);
            g.setFont(font);
            //获取文字长度
            int len = g.getFontMetrics(
                    g.getFont()).charsWidth(waterMark.toCharArray(),
                    0,
                    waterMark.length());
            int x = srcImgWidth - len - offsetX;
            int y = srcImgHeight - offsetY;
            g.drawString(waterMark, x, y);
            g.dispose();
            // 输出图片
            FileOutputStream outImgStream = new FileOutputStream(dstImgPath);
            ImageIO.write(bufImg, formatName, outImgStream);
            outImgStream.flush();
            outImgStream.close();
            logger.info(srcImgPath +" add water mark "+ waterMark + " OK!");
        } catch (Exception e) {
            e.printStackTrace();
            return -1;
        }
        return 0;
    }
    /**
     * 给图片添加水印、可设置水印图片旋转角度
     *
     * @param srcImgPath 源图片路径
     * @param dstImgPath 目标图片路径
     * @param waterMarkPath 水印图片路径 水印一般为gif或者png的，这样可设置透明度
     * @param degree 水印图片旋转角度, 360度
     * @param alpha 透明度 如0.5f
     * @param offsetX 偏移x default 20
     * @param offsetY 偏移y default 20
     * @param formatName 输出文件格式 JPG, JPEG, PNG, GIF, BMP
     * @return -1 添加失败，0 添加成功
     */
    public static int setWaterMarkFromImg(String srcImgPath, String dstImgPath,String waterMarkPath,
                                          Integer degree, float alpha, int offsetX, int offsetY,String formatName ) {
        OutputStream os = null;
        try {
            Image srcImg = ImageIO.read(new File(srcImgPath));

            BufferedImage buffImg = new BufferedImage(srcImg.getWidth(null), srcImg.getHeight(null),
                    BufferedImage.TYPE_INT_RGB);

            // 得到画笔对象
            Graphics2D g = buffImg.createGraphics();

            // 设置对线段的锯齿状边缘处理
            g.setRenderingHint(RenderingHints.KEY_INTERPOLATION, RenderingHints.VALUE_INTERPOLATION_BILINEAR);
            //srcImg.getScaledInstance(srcImg.getWidth(null), srcImg.getHeight(null), Image.SCALE_SMOOTH);
            g.drawImage(srcImg, 0,
                    0, null);

//            if (null != degree) {
//                // 设置水印旋转
//                g.rotate(Math.toRadians(degree), (double) buffImg.getWidth() / 2, (double) buffImg.getHeight() / 2);
//            }

            // 得到Image对象。
            Image imgWater = ImageIO.read(new File(waterMarkPath));

//            float alpha = 0.5f; // 透明度
            g.setComposite(AlphaComposite.getInstance(AlphaComposite.SRC_ATOP, alpha));

            // 表示水印图片的位置
            g.drawImage(imgWater, offsetX, offsetY, imgWater.getWidth(null),imgWater.getHeight(null),null);

            g.setComposite(AlphaComposite.getInstance(AlphaComposite.SRC_OVER));

            g.dispose();

            os = new FileOutputStream(dstImgPath);

            // 生成图片
            ImageIO.write(buffImg, formatName, os);

            logger.info(srcImgPath+" add water mark completed.");
        } catch (Exception e) {

            e.printStackTrace();
            return  -1;
        } finally {
            try {
                if (null != os)
                    os.close();
            } catch (Exception e) {
                e.printStackTrace();
                return -1;
            }
        }
        return 0;
    }


    /**
     * pdf设置文字水印
     *
     * @param inputPath
     * @param outPath
     * @param markStr
     * @param alpha 透明度如 0.2f
     * @param color 水印颜色
     * @throws DocumentException
     * @throws IOException
     */
    public static int setPdfWatermark(String inputPath, String outPath, String markStr, float alpha,
                                      BaseColor color)
            throws DocumentException, IOException {
        File file = new File(outPath);
        if (!file.exists()) {
            try {
                file.createNewFile();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
        BufferedOutputStream bufferOut = null;
        try {
            bufferOut = new BufferedOutputStream(new FileOutputStream(file));
        } catch (FileNotFoundException e1) {
            e1.printStackTrace();
            return -1;
        }

        PdfReader reader = new PdfReader(inputPath);
        PdfStamper stamper = new PdfStamper(reader, bufferOut);
        int total = reader.getNumberOfPages() + 1;
        PdfContentByte content;
        BaseFont base = BaseFont.createFont("STSong-Light", "UniGB-UCS2-H", BaseFont.EMBEDDED);
        // BaseFont base = BaseFont.createFont("/data/tmis/uploads/file/font/simsun.ttc,1", BaseFont.IDENTITY_H,
        // BaseFont.EMBEDDED);
        PdfGState gs = new PdfGState();
        for (int i = 1; i < total; i++) {
            content = stamper.getOverContent(i);// 在内容上方加水印
            //content = stamper.getUnderContent(i);// 在内容下方加水印
            gs.setFillOpacity(alpha);
            // content.setGState(gs);
            content.beginText();
            //content.setRGBColorFill(192, 192, 192);
            content.setColorFill(color);
            content.setFontAndSize(base, 50);
            content.setTextMatrix(100, 250);
            content.showTextAligned(Element.ALIGN_CENTER, markStr, 250, 400, 55);

            content.endText();
        }
        stamper.close();
        try {
            bufferOut.flush();
            bufferOut.close();
        } catch (IOException e) {
            e.printStackTrace();
            return -1;
        }
        return 0;
    }

    /**
     * word文字水印
     * @param srcPath 输入
     * @param dstPath 输出
     * @param markStr 水印文字
     * @param color 水印颜色 如 "#CC00FF"
     */
    public static int setWordWaterMark(String srcPath, String dstPath, String markStr,
                                        String color) throws FileNotFoundException,IOException,RuntimeException {
        // 获取传入的文件格式
        String fileType = ImageUtil.getFileType(srcPath);

        if(!fileType.equals("docx") && !fileType.equals("doc")){
            throw new RuntimeException(" file type not supported:"+fileType);
        }
        if ("docx".equals(fileType)) {
            File inputFile = new File(srcPath);
            XWPFDocument doc = null;
            try {
                doc = new XWPFDocument(new FileInputStream(inputFile));
            } catch (FileNotFoundException e) {
                e.printStackTrace();
            } catch (IOException e) {
                e.printStackTrace();
            }
            XWPFParagraph paragraph = doc.createParagraph();
            //XWPFRun run=paragraph.createRun();
            //run.setText("The Body:");
            // create header-footer
            XWPFHeaderFooterPolicy headerFooterPolicy = doc.getHeaderFooterPolicy();
            if (headerFooterPolicy == null) headerFooterPolicy = doc.createHeaderFooterPolicy();
            // create default Watermark - fill color black and not rotated
            headerFooterPolicy.createWatermark(markStr);
            // get the default header
            // Note: createWatermark also sets FIRST and EVEN headers
            // but this code does not updating those other headers
            XWPFHeader header = headerFooterPolicy.getHeader(XWPFHeaderFooterPolicy.DEFAULT);
            paragraph = header.getParagraphArray(0);
//            logger.debug(paragraph.getCTP().getRArray(0));
//            logger.debug(paragraph.getCTP().getRArray(0).getPictArray(0));
            // get com.microsoft.schemas.vml.CTShape where fill color and rotation is set
            org.apache.xmlbeans.XmlObject[] xmlobjects = paragraph.getCTP().getRArray(0).getPictArray(0).selectChildren(
                    new javax.xml.namespace.QName("urn:schemas-microsoft-com:vml", "shape"));
            if (xmlobjects.length > 0) {
                com.microsoft.schemas.vml.CTShape ctshape = (com.microsoft.schemas.vml.CTShape) xmlobjects[0];
                // set fill color
                ctshape.setFillcolor(color);
                // set rotation
                ctshape.setStyle(ctshape.getStyle() + ";rotation:315");

            }
            File file = new File(dstPath);
            if (!file.exists()) {
                try {
                    file.createNewFile();
                } catch (IOException e) {
                    e.printStackTrace();
                    return -1;
                }
            }
            try {
                doc.write(new FileOutputStream(file));
                doc.close();
            } catch (FileNotFoundException e) {
                e.printStackTrace();
                return -1;
            } catch (IOException e) {
                e.printStackTrace();
                return -1;
            }
        } else if ("doc".equals(fileType)) {
            logger.error("doc file not supported.");
            return -1;
        }

        return 0;
    }

    /*
     * 为Excel打上水印工具函数 请自行确保参数值，以保证水印图片之间不会覆盖。
     * 在计算水印的位置的时候，并没有考虑到单元格合并的情况，请注意
     *
     * @param wb
     *            Excel Workbook
     * @param sheet
     *            需要打水印的Excel
     * @param waterRemarkPath
     *            水印地址，classPath，目前只支持png格式的图片，
     *            因为非png格式的图片打到Excel上后可能会有图片变红的问题，且不容易做出透明效果。
     *            同时请注意传入的地址格式，应该为类似："\\excelTemplate\\test.png"
     * @param startXCol
     *            水印起始列
     * @param startYRow
     *            水印起始行
     * @param betweenXCol
     *            水印横向之间间隔多少列
     * @param betweenYRow
     *            水印纵向之间间隔多少行
     * @param XCount
     *            横向共有水印多少个
     * @param YCount
     *            纵向共有水印多少个
     * @param waterRemarkWidth
     *            水印图片宽度为多少列
     * @param waterRemarkHeight
     *            水印图片高度为多少行
     * @throws IOException
     */
    public static void putWaterRemarkToExcel(Workbook wb, Sheet sheet, String waterRemarkPath, int startXCol,
                                             int startYRow, int betweenXCol, int betweenYRow, int XCount, int YCount, int waterRemarkWidth,
                                             int waterRemarkHeight) throws IOException {

        // 校验传入的水印图片格式
        String fileType = ImageUtil.getFileType(waterRemarkPath);
        if (!fileType.equals("png")) {
            throw new RuntimeException("Excel only support png water mark.");
        }

        // 加载图片
        ByteArrayOutputStream byteArrayOut = new ByteArrayOutputStream();
        InputStream imageIn = new FileInputStream(waterRemarkPath);
        if (null == imageIn || imageIn.available() < 1) {
            throw new RuntimeException("add water mark to Excel,failed to read water mark image, file not exist.");
        }
        BufferedImage bufferImg = ImageIO.read(imageIn);
        if (null == bufferImg) {
            throw new RuntimeException("add water mark to Excel,failed to read water mark image.");
        }
        ImageIO.write(bufferImg, "png", byteArrayOut);

        // 开始打水印
        Drawing drawing = sheet.createDrawingPatriarch();

        // 按照共需打印多少行水印进行循环
        for (int yCount = 0; yCount < YCount; yCount++) {
            // 按照每行需要打印多少个水印进行循环
            for (int xCount = 0; xCount < XCount; xCount++) {
                // 创建水印图片位置
                int xIndexInteger = startXCol + (xCount * waterRemarkWidth) + (xCount * betweenXCol);
                int yIndexInteger = startYRow + (yCount * waterRemarkHeight) + (yCount * betweenYRow);
                /*
                 * 参数定义： 第一个参数是（x轴的开始节点）； 第二个参数是（是y轴的开始节点）； 第三个参数是（是x轴的结束节点）；
                 * 第四个参数是（是y轴的结束节点）； 第五个参数是（是从Excel的第几列开始插入图片，从0开始计数）；
                 * 第六个参数是（是从excel的第几行开始插入图片，从0开始计数）； 第七个参数是（图片宽度，共多少列）；
                 * 第8个参数是（图片高度，共多少行）；
                 */
                ClientAnchor anchor = drawing.createAnchor(0, 0, 0, 0, xIndexInteger,
                        yIndexInteger, xIndexInteger + waterRemarkWidth, yIndexInteger + waterRemarkHeight);

                Picture pic = drawing.createPicture(anchor,
                        wb.addPicture(byteArrayOut.toByteArray(), Workbook.PICTURE_TYPE_PNG));
                pic.resize();
            }
        }

    }

    /**
     * excel设置水印
     * @param srcPath
     * @param dstPath
     * @param markStr
     */
    public static void setExcelWaterMark(String srcPath, String dstPath, String markStr, String imgTempPath) throws Exception {
        // 获取传入的文件格式
        String fileType = ImageUtil.getFileType(srcPath);

        if(!fileType.equals("xlsx") && !fileType.equals("xls")){
            throw new RuntimeException(" file type not supported:"+fileType);
        }
        //读取excel文件
        Workbook wb = null;
        try {
            wb = new XSSFWorkbook(new FileInputStream(srcPath));
        } catch (FileNotFoundException e1) {
            e1.printStackTrace();
        } catch (IOException e1) {
            e1.printStackTrace();
        }
        //设置水印图片路径
        String imgPath = imgTempPath + "/waterImage.png";
        try {
            ImageUtil.createImageFromTxt(markStr, imgPath, 300,300, Color.white, 50,"黑体","png");
        } catch (IOException e1) {
            e1.printStackTrace();
        }
        //获取excel sheet个数
        int sheets = wb.getNumberOfSheets();
        //循环sheet给每个sheet添加水印
        for (int i = 0; i < sheets; i++) {
            Sheet sheet = wb.getSheetAt(i);
            //excel加密只读
            //sheet.protectSheet(UUID.randomUUID().toString());
            //获取excel实际所占行
            //int row = sheet.getFirstRowNum() + sheet.getLastRowNum();
            int row = 0;
            //获取excel实际所占列
            int cell = 0;
            /*
             * if(null != sheet.getRow(sheet.getFirstRowNum())) {
             * cell = sheet.getRow(sheet.getFirstRowNum()).getLastCellNum() + 1;
             * }
             */

            //根据行与列计算实际所需多少水印
            try {
                putWaterRemarkToExcel(wb, sheet, imgPath, 0, 0, 5, 5, cell / 5 + 1, row / 5 + 1, 0, 0);
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
        ByteArrayOutputStream os = new ByteArrayOutputStream();
        try {
            wb.write(os);
            wb.close();
        } catch (IOException e) {
            e.printStackTrace();
        }

        byte[] content = os.toByteArray();
        // Excel文件生成后存储的位置。
        File file = new File(dstPath);
        OutputStream fos = null;
        try {
            fos = new FileOutputStream(file);
            fos.write(content);
            os.close();
            fos.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
        File imageTempFile = new File(imgPath);
        if(imageTempFile.exists()) {
            imageTempFile.delete();
        }
    }


    //改变所有文本，不改变样式
    public static void setPPTWaterMark(String srcPath,String targetPath, String markStr, String imgTempPath) throws IOException ,RuntimeException{
        // 获取传入的文件格式
        String fileType = ImageUtil.getFileType(srcPath);

        if(!fileType.equals("pptx") && !fileType.equals("ppt")){
            throw new RuntimeException(" file type not supported:"+fileType);
        }
        //设置水印图片路径
        String imgPath =  imgTempPath + "/waterImage.png";
        try {
            ImageUtil.createImageFromTxt(markStr, imgPath, 300,300, Color.white, 50,"黑体","png");

        } catch (IOException e1) {
            e1.printStackTrace();
        }
        if("pptx".equals(fileType)) {
            XMLSlideShow slideShow = new XMLSlideShow(new FileInputStream(srcPath));
            byte[] pictureData = IOUtils.toByteArray(new FileInputStream(imgPath));
//            byte[] pictureData = IOUtils.toByteArray(new FileInputStream("E:\\1.png"));
            PictureData pictureData1 = slideShow.addPicture(pictureData, PictureType.PNG);
            for (XSLFSlide slide : slideShow.getSlides()) {
                XSLFPictureShape pictureShape = slide.createPicture(pictureData1);
                pictureShape.setAnchor(new Rectangle(50, 300, 100, 100));
                CTSlide ctSlide = slide.getXmlObject();
                XmlObject[] allText = ctSlide.selectPath(
                        "declare namespace a='http://schemas.openxmlformats.org/drawingml/2006/main' " +
                                ".//a:t"
                );
                /*
                 * for (int i = 0; i < allText.length; i++) {
                 * if (allText[i] instanceof XmlString) {
                 * XmlString xmlString = (XmlString)allText[i];
                 * String text = xmlString.getStringValue();
                 * if (text==null||text.equals("")) continue;
                 * if (status==1) xmlString.setStringValue(WordAPi.Encrypt(text));
                 * else xmlString.setStringValue(WordAPi.Decrypt(text));
                 * }
                 * }
                 */
            }

            FileOutputStream out = new FileOutputStream(targetPath);
            slideShow.write(out);
            slideShow.close();
            out.close();
        } else if("ppt".equals(fileType)) {
            logger.error("ppt not supported.");
            /*
             * HSLFSlideShow ppt = new HSLFSlideShow();
             * SlideShow ppt = new SlideShow(new HSLFSlideShow("PPT测试.ppt"));
             * SlideShow slideShow = new (new FileInputStream(path));
             * byte[] pictureData = IOUtils.toByteArray(new FileInputStream(imgPath));
             * // byte[] pictureData = IOUtils.toByteArray(new FileInputStream("E:\\1.png"));
             * PictureData pictureData1 = slideShow.addPicture(pictureData, PictureType.PNG);
             * for (XSLFSlide slide : slideShow.getSlides()) {
             * XSLFPictureShape pictureShape = slide.createPicture(pictureData1);
             * pictureShape.setAnchor(new java.awt.Rectangle(50, 300, 100, 100));
             * CTSlide ctSlide = slide.getXmlObject();
             * XmlObject[] allText = ctSlide.selectPath(
             * "declare namespace a='http://schemas.openxmlformats.org/drawingml/2006/main' " +
             * ".//a:t"
             * );
             * }
             *
             * FileOutputStream out = new FileOutputStream(targetpath);
             * slideShow.write(out);
             * slideShow.close();
             * out.close();
             */
        }
    }
    public static void main(String[] args) {


        setWaterMarkFromImg("c:/zhh/id.png","c:/zhh/idimg2.png",
                "c:/zhh/mywater.png",30,0.1f,30,30,"png");
    }
}

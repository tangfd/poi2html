package com.tfd.pdf2html;

import com.tfd.POIUtils;
import org.apache.commons.io.FileUtils;
import org.apache.commons.io.IOUtils;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.rendering.PDFRenderer;

import javax.imageio.ImageIO;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

/**
 * pdf 转 image
 *
 * @author TangFD@HF 2018/5/29
 */
public class PDF2Image {

    public static void pdf2Image(File pdfFile, String htmlDir, String imgDir, String imgWebPath, String password) throws IOException {
        POIUtils.createDir(imgDir);
        StringBuilder content = new StringBuilder();
        content.append("<!doctype html>");
        content.append("<head>");
        content.append("<meta charset=\"UTF-8\">");
        content.append("<style>");
        content.append(".image {background-color:#fff; text-align:center; width:100%; max-width:100%;margin-top:6px;}");
        content.append("</style>");
        content.append("</head>");
        content.append("<body style=\"background-color:gray;\">");
        content.append("<div class='image'>");

        PDDocument document = PDDocument.load(pdfFile, password);
        int pages = document.getNumberOfPages();
        PDFRenderer reader = new PDFRenderer(document);
        //遍历处理pdf
        for (int i = 0; i < pages; i++) {
            FileOutputStream out = null;
            try {
//                    BufferedImage image = reader.renderImageWithDPI(i, 130, ImageType.RGB);
                BufferedImage image = reader.renderImage(i, 1.5f);
                //生成图片,保存位置
                String img = POIUtils.getUUID() + ".png";
                out = new FileOutputStream(imgDir + img);
                ImageIO.write(image, "png", out);
                //将图片路径追加到网页文件里
                content.append("<img src=\"").append(imgWebPath).append(img).append("\"/><br>");
                out.flush();
            } finally {
                IOUtils.closeQuietly(out);
            }
        }

        document.close();
        content.append("</div>");
        content.append("</body></html>");
        //生成网页文件
        String targetFileName = htmlDir + pdfFile.getName() + ".html";
        FileUtils.writeStringToFile(new File(targetFileName), content.toString(), "utf-8");
    }
}

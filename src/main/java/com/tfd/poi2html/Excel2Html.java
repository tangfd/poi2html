package com.tfd.poi2html;

import com.tfd.POIUtils;
import com.tfd.xlsx.XssfExcelToHtmlConverter;
import org.apache.commons.io.FileUtils;
import org.apache.poi.hssf.converter.ExcelToHtmlConverter;
import org.apache.poi.hssf.usermodel.HSSFPictureData;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.PictureData;
import org.w3c.dom.Document;

import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.transform.OutputKeys;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;
import java.io.*;
import java.util.ArrayList;
import java.util.List;

/**
 * excel转html
 *
 * @author TangFD@HF 2018/5/28
 */
public class Excel2Html {
    public static void xls2Html(File file, String htmlDir, String imgDir, String imgWebPath) throws Exception {
        if (file.getName().endsWith("xlsx")) {
            xlsx2Html(file, htmlDir, imgDir, imgWebPath);
        } else if (file.getName().endsWith("xls")) {
            xls2html(file, htmlDir, imgDir, imgWebPath);
        } else {
            throw new RuntimeException("file not xls!");
        }
    }

    private static void xlsx2Html(File xlsFile, String htmlDir, String imgDir, String imgWebPath) throws Exception {
        Document doc = XssfExcelToHtmlConverter.process(xlsFile);
        DOMSource domSource = new DOMSource(doc);
        ByteArrayOutputStream outStream = new ByteArrayOutputStream();
        StreamResult streamResult = new StreamResult(outStream);
        TransformerFactory tf = TransformerFactory.newInstance();
        Transformer serializer = tf.newTransformer();
        serializer.setOutputProperty(OutputKeys.ENCODING, "UTF-8");
        serializer.setOutputProperty(OutputKeys.INDENT, "no");
        serializer.setOutputProperty(OutputKeys.METHOD, "html");
        serializer.transform(domSource, streamResult);
        outStream.close();

        String content = new String(outStream.toByteArray());
        content = content.replace("</style>", "tr,td{border: 1px solid;}</style>");
        Object pictureDatas = doc.getUserData("xlsx-pics");
        if (pictureDatas != null) {
            String imgContent = getImgHtmlContent(imgDir, imgWebPath, (List<PictureData>) pictureDatas);
            content = content.replace("</body>", imgContent);
        }

        String targetFileName = htmlDir + xlsFile.getName() + ".html";
        FileUtils.writeStringToFile(new File(targetFileName), content, "utf-8");
    }

    private static String getImgHtmlContent(String imgDir, String imgWebPath,
                                            List<PictureData> pictureDatas) throws IOException {
        StringBuilder builder = new StringBuilder();
        if (pictureDatas == null || pictureDatas.size() == 0) {
            return "</body>";
        }

        for (PictureData pictureData : pictureDatas) {
            try {
                String imgName = POIUtils.getUUID() + "." + pictureData.suggestFileExtension();
                builder.append("<img src=\"").append(imgWebPath).append(imgName).append("\" style=\"width:7.5in;height:4.5in;vertical-align:text-bottom;\"></p>");
                FileOutputStream out = new FileOutputStream(imgDir + imgName);
                out.write(pictureData.getData());
                out.close();
            } catch (FileNotFoundException e) {
                e.printStackTrace();
            }
        }

        builder.append("</body>");
        return builder.toString();
    }

    private static void xls2html(File xlsFile, String htmlDir, String imgDir, String imgWebPath) throws Exception {
        InputStream input = new FileInputStream(xlsFile);
        HSSFWorkbook excelBook = new HSSFWorkbook(input);
        Document newDocument = DocumentBuilderFactory.newInstance().newDocumentBuilder().newDocument();
        ExcelToHtmlConverter excelToHtmlConverter = new ExcelToHtmlConverter(newDocument);
        excelToHtmlConverter.processWorkbook(excelBook);
        Document htmlDocument = excelToHtmlConverter.getDocument();
        ByteArrayOutputStream outStream = new ByteArrayOutputStream();
        DOMSource domSource = new DOMSource(htmlDocument);
        StreamResult streamResult = new StreamResult(outStream);
        TransformerFactory tf = TransformerFactory.newInstance();
        Transformer serializer = tf.newTransformer();
        serializer.setOutputProperty(OutputKeys.ENCODING, "utf-8");
        serializer.setOutputProperty(OutputKeys.INDENT, "yes");
        serializer.setOutputProperty(OutputKeys.METHOD, "html");
        serializer.transform(domSource, streamResult);
        outStream.close();

        String content = new String(outStream.toByteArray());
        content = content.replace("</style>", "tr,td{border: 1px solid;}</style>");
        List<HSSFPictureData> pictures = excelBook.getAllPictures();
        if (pictures != null) {
            String imgContent = getImgHtmlContent(imgDir, imgWebPath, new ArrayList<PictureData>(pictures));
            content = content.replace("</body>", imgContent);
        }

        String targetFilePath = htmlDir + xlsFile.getName() + ".html";
        FileUtils.writeStringToFile(new File(targetFilePath), content, "utf-8");
    }
}

package com.tfd.poi2html;

import com.tfd.POIUtils;
import org.apache.commons.io.FileUtils;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.converter.PicturesManager;
import org.apache.poi.hwpf.converter.WordToHtmlConverter;
import org.apache.poi.hwpf.usermodel.Picture;
import org.apache.poi.hwpf.usermodel.PictureType;
import org.apache.poi.xwpf.converter.core.BasicURIResolver;
import org.apache.poi.xwpf.converter.core.FileImageExtractor;
import org.apache.poi.xwpf.converter.xhtml.XHTMLConverter;
import org.apache.poi.xwpf.converter.xhtml.XHTMLOptions;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.w3c.dom.Document;

import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.transform.OutputKeys;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;
import java.io.*;
import java.util.List;

/**
 * word转html
 *
 * @author TangFD@HF 2018/5/28
 */
public class Word2Html {

    public static void word2Html(File file, String htmlDir, String imgDir, String imgWebPath) throws Exception {
        if (file.getName().endsWith("docx")) {
            docx2Html(file, htmlDir, imgDir, imgWebPath);
        } else if (file.getName().endsWith("doc")) {
            doc2Html(file, htmlDir, imgDir, imgWebPath);
        } else {
            throw new RuntimeException("file not doc!");
        }
    }

    private static void docx2Html(File docFile, String htmlDir, String imgDir, String imgWebPath) throws IOException {
        XWPFDocument document = new XWPFDocument(new FileInputStream(docFile));
        XHTMLOptions options = XHTMLOptions.create();
        // 存放图片的文件夹
        String uuid = POIUtils.getUUID();
        imgDir = imgDir + uuid + File.separator;
        POIUtils.createDir(imgDir);
        options.setExtractor(new FileImageExtractor(new File(imgDir)));
        // html中图片的路径
        options.URIResolver(new BasicURIResolver(imgWebPath + uuid + File.separator));
        String targetFileName = htmlDir + docFile.getName() + ".html";
        OutputStreamWriter outputStreamWriter = new OutputStreamWriter(new FileOutputStream(targetFileName), "utf-8");
        XHTMLConverter xhtmlConverter = (XHTMLConverter) XHTMLConverter.getInstance();
        xhtmlConverter.convert(document, outputStreamWriter, options);
    }

    private static void doc2Html(File docFile, String htmlDir, String imgDir, final String imgWebPath) throws Exception {
        InputStream input = new FileInputStream(docFile);
        HWPFDocument wordDocument = new HWPFDocument(input);
        Document newDocument = DocumentBuilderFactory.newInstance().newDocumentBuilder().newDocument();
        WordToHtmlConverter wordToHtmlConverter = new WordToHtmlConverter(newDocument);
        wordToHtmlConverter.setPicturesManager(new PicturesManager() {
            public String savePicture(byte[] bytes, PictureType pictureType, String s, float v, float v1) {
                return s;
            }
        });
        wordToHtmlConverter.processDocument(wordDocument);
        List pics = wordDocument.getPicturesTable().getAllPictures();
        String uuid = POIUtils.getUUID();
        imgDir = imgDir + uuid + File.separator;
        POIUtils.createDir(imgDir);
        if (pics != null) {
            for (Object pic1 : pics) {
                Picture pic = (Picture) pic1;
                try {
                    pic.writeImageContent(new FileOutputStream(imgDir + pic.suggestFullFileName()));
                } catch (FileNotFoundException e) {
                    e.printStackTrace();
                }
            }
        }
        Document document = wordToHtmlConverter.getDocument();
        ByteArrayOutputStream outStream = new ByteArrayOutputStream();
        DOMSource domSource = new DOMSource(document);
        StreamResult streamResult = new StreamResult(outStream);
        TransformerFactory tf = TransformerFactory.newInstance();
        Transformer serializer = tf.newTransformer();
        serializer.setOutputProperty(OutputKeys.ENCODING, "utf-8");
        serializer.setOutputProperty(OutputKeys.INDENT, "yes");
        serializer.setOutputProperty(OutputKeys.METHOD, "html");
        serializer.transform(domSource, streamResult);
        outStream.close();
        String content = new String(outStream.toByteArray());
        content = content.replace("<img src=\"", "<img src=\"" + imgWebPath + uuid + "/");
        String targetFilePath = htmlDir + docFile.getName() + ".html";
        FileUtils.writeStringToFile(new File(targetFilePath), content, "utf-8");
    }
}

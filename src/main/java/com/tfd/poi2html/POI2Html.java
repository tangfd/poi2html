package com.tfd.poi2html;

import com.tfd.POIUtils;
import com.tfd.pdf2html.PDF2Image;

import java.io.File;
import java.io.IOException;

/**
 * @author TangFD@HF 2018/5/29
 */
public class POI2Html {
    public static void main(String[] args) throws Exception {
        poi2Html("xls/test.xls", "xls", "xls/img", "img");
        poi2Html("xls/test.xlsx", "xls", "xls/img", "img");
    }

    public static void pdf2HtmlByPassword(String filePath, String htmlDir, String imgDir, String imgWebPath, String password) throws IOException {
        File file = POIUtils.checkFileExists(filePath);
        htmlDir = POIUtils.dealTargetDir(htmlDir);
        imgDir = POIUtils.dealTargetDir(imgDir);
        imgWebPath = POIUtils.dealTargetDir(imgWebPath);
        PDF2Image.pdf2Image(file, htmlDir, imgDir, imgWebPath, password);
    }

    /**
     * @param filePath   待转换的原文件路径
     * @param htmlDir    生成的html文件存放目录
     * @param imgDir     原文件的图片存放目录
     * @param imgWebPath web应用访问图片的系统路径
     * @throws Exception
     */
    public static void poi2Html(String filePath, String htmlDir, String imgDir, String imgWebPath) throws Exception {
        File file = POIUtils.checkFileExists(filePath);
        htmlDir = POIUtils.dealTargetDir(htmlDir);
        imgDir = POIUtils.dealTargetDir(imgDir);
        imgWebPath = POIUtils.dealTargetDir(imgWebPath);
        if (filePath.endsWith("ppt") || filePath.endsWith("pptx")) {
            PPT2Image.ppt2Html(file, htmlDir, imgDir, imgWebPath);
        } else if (filePath.endsWith("doc") || filePath.endsWith("docx")) {
            Word2Html.word2Html(file, htmlDir, imgDir, imgWebPath);
        } else if (filePath.endsWith("xls") || filePath.endsWith("xlsx")) {
            Excel2Html.xls2Html(file, htmlDir, imgDir, imgWebPath);
        } else if (filePath.endsWith("pdf")) {
            PDF2Image.pdf2Image(file, htmlDir, imgDir, imgWebPath, null);
        } else {
            throw new RuntimeException("invalid file type![filePath : " + filePath + "]");
        }
    }

    /**
     * @param file       待转换的原文件
     * @param htmlDir    生成的html文件存放目录
     * @param imgDir     原文件的图片存放目录
     * @param imgWebPath web应用访问图片的系统路径
     * @throws Exception
     */
    public static void poi2Html(File file, String htmlDir, String imgDir, String imgWebPath) throws Exception {
        String filePath = file.getPath();
        if (!file.exists()) {
            throw new RuntimeException("file not exists ![filepath = " + filePath + "]");
        }

        htmlDir = POIUtils.dealTargetDir(htmlDir);
        imgDir = POIUtils.dealTargetDir(imgDir);
        imgWebPath = POIUtils.dealTargetDir(imgWebPath);
        poi2Html(filePath, htmlDir, imgDir, imgWebPath);
    }
}

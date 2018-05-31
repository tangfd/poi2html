package com.tfd.xlsx;

import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.util.Beta;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

/**
 * Common class for {@link XssfExcelToHtmlUtils}
 *
 * @author Sergey Vladimirov (vlsergey {at} gmail {dot} com)
 * @author wenfeng.xu wechat id :italybaby
 */
@Beta
public class AbstractXssfExcelUtils {
    public static final String EMPTY = "";
    private static final short EXCEL_COLUMN_WIDTH_FACTOR = 256;
    private static final int UNIT_OFFSET_LENGTH = 7;

    public static String getAlign(short alignment) {
        switch (alignment) {
            case XSSFCellStyle.ALIGN_CENTER:
                return "center";
            case XSSFCellStyle.ALIGN_CENTER_SELECTION:
                return "center";
            case XSSFCellStyle.ALIGN_FILL:
                // XXX: shall we support fill?
                return "";
            case XSSFCellStyle.ALIGN_GENERAL:
                return "";
            case XSSFCellStyle.ALIGN_JUSTIFY:
                return "justify";
            case XSSFCellStyle.ALIGN_LEFT:
                return "left";
            case XSSFCellStyle.ALIGN_RIGHT:
                return "right";
            default:
                return "";
        }
    }

    public static String getBorderStyle(short xlsBorder) {
        final String borderStyle;
        switch (xlsBorder) {
            case XSSFCellStyle.BORDER_NONE:
                borderStyle = "none";
                break;
            case XSSFCellStyle.BORDER_DASH_DOT:
            case XSSFCellStyle.BORDER_DASH_DOT_DOT:
            case XSSFCellStyle.BORDER_DOTTED:
            case XSSFCellStyle.BORDER_HAIR:
            case XSSFCellStyle.BORDER_MEDIUM_DASH_DOT:
            case XSSFCellStyle.BORDER_MEDIUM_DASH_DOT_DOT:
            case XSSFCellStyle.BORDER_SLANTED_DASH_DOT:
                borderStyle = "dotted";
                break;
            case XSSFCellStyle.BORDER_DASHED:
            case XSSFCellStyle.BORDER_MEDIUM_DASHED:
                borderStyle = "dashed";
                break;
            case XSSFCellStyle.BORDER_DOUBLE:
                borderStyle = "double";
                break;
            default:
                borderStyle = "solid";
                break;
        }
        return borderStyle;
    }

    public static String getBorderWidth(short xlsBorder) {
        final String borderWidth;
        switch (xlsBorder) {
            case XSSFCellStyle.BORDER_MEDIUM_DASH_DOT:
            case XSSFCellStyle.BORDER_MEDIUM_DASH_DOT_DOT:
            case XSSFCellStyle.BORDER_MEDIUM_DASHED:
                borderWidth = "2pt";
                break;
            case XSSFCellStyle.BORDER_THICK:
                borderWidth = "thick";
                break;
            default:
                borderWidth = "thin";
                break;
        }
        return borderWidth;
    }

    /**
     * 转换为网页颜色
     *
     * @param color
     * @return
     */
    public static String getColor(XSSFColor color) {
        String result = color.getARGBHex();
        if (result != null) {
            result = "#" + result.substring(2, result.length());
        } else {
            result = "";
        }
        return result;
    }

    /**
     * See <a href=
     * "http://apache-poi.1045710.n5.nabble.com/Excel-Column-Width-Unit-Converter-pixels-excel-column-width-units-td2301481.html"
     * >here</a> for Xio explanation and details
     */
    public static int getColumnWidthInPx(int widthUnits) {
        int pixels = (widthUnits / EXCEL_COLUMN_WIDTH_FACTOR)
                * UNIT_OFFSET_LENGTH;

        int offsetWidthUnits = widthUnits % EXCEL_COLUMN_WIDTH_FACTOR;
        pixels += Math.round(offsetWidthUnits
                / ((float) EXCEL_COLUMN_WIDTH_FACTOR / UNIT_OFFSET_LENGTH));

        return pixels;
    }

    /**
     * @param mergedRanges map of sheet merged ranges built with
     *                     {@link XssfExcelToHtmlUtils#buildMergedRangesMap(HSSFSheet)}
     * @return {@link CellRangeAddress} from map if cell with specified row and
     * column numbers contained in found range, <tt>null</tt> otherwise
     */
    public static CellRangeAddress getMergedRange(
            CellRangeAddress[][] mergedRanges, int rowNumber, int columnNumber) {
        CellRangeAddress[] mergedRangeRowInfo = rowNumber < mergedRanges.length ? mergedRanges[rowNumber]
                : null;
        CellRangeAddress cellRangeAddress = mergedRangeRowInfo != null
                && columnNumber < mergedRangeRowInfo.length ? mergedRangeRowInfo[columnNumber]
                : null;

        return cellRangeAddress;
    }

    public static boolean isEmpty(String str) {
        return str == null || str.length() == 0;
    }

    public static boolean isNotEmpty(String str) {
        return !isEmpty(str);
    }

    public static XSSFWorkbook loadXls(File xlsFile) throws IOException {
        final FileInputStream inputStream = new FileInputStream(xlsFile);
        try {
            return new XSSFWorkbook(inputStream);
        } finally {
            IOUtils.closeQuietly(inputStream);
        }
    }


}


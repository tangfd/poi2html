package com.tfd.xlsx;

import org.apache.poi.ss.usermodel.DataFormatter;

import java.util.Locale;

/**
 * @author wenfeng.xu wechat id :italybaby
 */
public class XssfDataFormatter extends DataFormatter {

    public XssfDataFormatter(Locale locale) {
        super(locale);

    }

    public XssfDataFormatter() {
        this(Locale.getDefault());
    }


}

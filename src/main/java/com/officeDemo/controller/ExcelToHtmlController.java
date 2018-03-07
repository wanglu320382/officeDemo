package com.officeDemo.controller;

import cn.afterturn.easypoi.cache.manager.POICacheManager;
import cn.afterturn.easypoi.excel.ExcelXorHtmlUtil;
import cn.afterturn.easypoi.excel.entity.ExcelToHtmlParams;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;

/**
 * Create by liulinhui on 2018/3/4 19:46
 */
@RequestMapping("/excelToHtml")
@Controller
public class ExcelToHtmlController {

    /**
     * 03 版本EXCEL预览
     */
    @RequestMapping("/03")
    public void toHtmlOf03Base(HttpServletResponse response) throws IOException, InvalidFormatException {
        ExcelToHtmlParams params = new ExcelToHtmlParams(WorkbookFactory.create(POICacheManager.getFile("excelToHtml/testExportTitleExcel.xls")));
        response.getOutputStream().write(ExcelXorHtmlUtil.excelToHtml(params).getBytes("gbk"));
    }

    /**
     * 03 版本EXCEL预览
     */
    @RequestMapping("/08")
    public void toHtmlOf08Base(HttpServletResponse response) throws IOException, InvalidFormatException {
        ExcelToHtmlParams params = new ExcelToHtmlParams(WorkbookFactory.create(POICacheManager.getFile("excelToHtml/08.xls")));
        response.getOutputStream().write(ExcelXorHtmlUtil.excelToHtml(params).getBytes("gbk"));
    }
    /**
     * 07 版本EXCEL预览
     */
    @RequestMapping("/07")
    public void toHtmlOf07Base(HttpServletResponse response) throws IOException, InvalidFormatException {
        ExcelToHtmlParams params = new ExcelToHtmlParams(WorkbookFactory.create(POICacheManager.getFile("excelToHtml/testExportTitleExcel.xlsx")));
        response.getOutputStream().write(ExcelXorHtmlUtil.excelToHtml(params).getBytes("gbk"));
    }
    /**
     * 03 版本EXCEL预览
     */
    @RequestMapping("/03img")
    public void toHtmlOf03Img(HttpServletResponse response) throws IOException, InvalidFormatException {
        ExcelToHtmlParams params = new ExcelToHtmlParams(WorkbookFactory.create(POICacheManager.getFile("excelToHtml/exporttemp_img.xls")),true,"yes");
        response.getOutputStream().write(ExcelXorHtmlUtil.excelToHtml(params).getBytes());
    }
    /**
     * 03 版本EXCEL预览
     */
    @RequestMapping("/20180128")
    public void toHtmlOf20180128Base(HttpServletResponse response) throws IOException, InvalidFormatException {
        ExcelToHtmlParams params = new ExcelToHtmlParams(WorkbookFactory.create(POICacheManager.getFile("excelToHtml/20180128.xls")),true,"yes");
        response.getOutputStream().write(ExcelXorHtmlUtil.excelToHtml(params).getBytes());
    }

    /**
     * 03 版本EXCEL预览
     */
    @RequestMapping("/20180129")
    public void toHtmlOf20180129Base(HttpServletResponse response) throws IOException, InvalidFormatException {
        ExcelToHtmlParams params = new ExcelToHtmlParams(WorkbookFactory.create(POICacheManager.getFile("excelToHtml/20180129.xlsx")),true,"yes");
        response.getOutputStream().write(ExcelXorHtmlUtil.excelToHtml(params).getBytes());
    }
}

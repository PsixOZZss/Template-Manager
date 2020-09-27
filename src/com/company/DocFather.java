package com.company;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFHeader;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPageMar;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSectPr;

import java.math.BigInteger;
import java.util.List;

public class DocFather {
    XWPFDocument docxModel;
    XWPFHeader header;

    List<String> fieldName;
    CTSectPr sectPr;
    CTPageMar pageMar;

    private DocFather(List<String>fieldName){
        this.fieldName=fieldName;
        createDocx();
    }
    private void createDocx(){
        docxModel = new XWPFDocument();
        header = new XWPFHeader();
        header.setXWPFDocument(docxModel);
        sectPr = docxModel.getDocument().getBody().addNewSectPr();
        pageMar = sectPr.addNewPgMar();

    }
    private void setLeft(BigInteger value){
        pageMar.setLeft(value);
    }
    private void setRight(BigInteger value){
        pageMar.setRight(value);
    }
    private void setTop(BigInteger value){
        pageMar.setTop(value);
    }
    private void setBot(BigInteger value){
        pageMar.setBottom(value);
    }
    public List<String>getFieldName(){
        return fieldName;
    }
    public void putNSave(List<String>fieldText){


    }

    private XWPFTable newTable(int rows, int cols){
        XWPFTable table = docxModel.createTable(rows, cols);
        return table;
    }
}

package com.company;

import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTP;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTR;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSectPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTText;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPageMar;

import java.io.FileOutputStream;
import java.math.BigInteger;

public class doc1{
    public doc1(){
        createDocx();
    }
    private void createDocx(){
        try {
            XWPFDocument docxModel = new XWPFDocument();
            XWPFHeader header = new XWPFHeader();
            header.setXWPFDocument(docxModel);

            CTSectPr sectPr = docxModel.getDocument().getBody().addNewSectPr();
            CTPageMar pageMar = sectPr.addNewPgMar();
            pageMar.setLeft(BigInteger.valueOf(570L));
            pageMar.setTop(BigInteger.valueOf(200L));
            pageMar.setRight(BigInteger.valueOf(570L));
            pageMar.setBottom(BigInteger.valueOf(200L));
            createTextWithoutSpacing(docxModel, ParagraphAlignment.CENTER, "Дата подачи заявки 27.07.2020 г.", 10, false);
            createTextWithoutSpacing(docxModel, ParagraphAlignment.CENTER, "Заявка на перевозку груза", 16, true);
            createText(docxModel, ParagraphAlignment.CENTER, "", 14, false);
            createTextWithoutSpacing(docxModel, ParagraphAlignment.LEFT, "ЗАКАЗЧИК:  ", 11, true, false, "ООО  «Союз Транс»", 11, true, false);
            createText(docxModel, ParagraphAlignment.LEFT, "Адрес:  445054  РФ  Самарская область, г. Тольятти, ул. Мира, 93-12,  тел./факс: +79171260468 Владислав", 9, false);
            createTextWithoutSpacing(docxModel, ParagraphAlignment.LEFT, "ИСПОЛНИТЕЛЬ:  ", 11, true, false, "ИП Фокин Александр Борисович", 11, true, true);
            createText(docxModel, ParagraphAlignment.LEFT, "Адрес: 423812, РТ, г. Набережные Челны, пр-кт Р. Беляева, д.29, кв. 65, тел. 8/85557/7-04-07", 9, false);
            XWPFTable table = docxModel.createTable(20, 2);
            tableText(table, 0, 0, 10, true, "Маршрут");
            tableText(table, 0, 1, 10, false, "М.О.Дмистровский р-н  Рогачева -Екатеринбург");
            tableText(table, 1, 0, 10, true, "Наименование груза, вес, тип загрузки");
            tableText(table, 1, 1, 10, false, "Плитка на паллетах и декоративный камень.");
            tableText(table, 2, 0, 10, true, "1. Дата и время подачи под погрузку");
            tableText(table, 2, 1, 10, false, "27.07.2020  до15:00");
            tableText(table, 3, 0, 10, false, "Адрес  места погрузки");
            tableText(table, 3, 1, 10, false, "Дмитровский р-н.Рогачево Советская ул.Вл36 ООО.Монолитсрой");
            tableText(table, 4, 0, 10, false, "Грузоотправитель, контактное лицо");
            tableText(table, 4, 1, 10, false, "8(969)010 93 90 Мария");
            tableText(table, 5, 0, 10, true, "2. Дата и время подачи под погрузку");
            tableText(table, 5, 1, 10, false, "-");
            tableText(table, 6, 0, 10, false, "Адрес второго места погрузки");
            tableText(table, 6, 1, 10, false, "-");
            tableText(table, 7, 0, 10, false, "Грузоотправитель, контактное лицо");
            tableText(table, 7, 1, 10, false, "-");
            tableText(table, 8, 0, 10, true, "1. Дата и время подачи под выгрузку");
            tableText(table, 8, 1, 10, false, "30.07.2020 к 9-00 обязательно позвонить и предупредить о приезде..");
            tableText(table, 9, 0, 10, false, "Адрес первого места выгрузки");
            tableText(table, 9, 1, 10, false, "Свердловская об. Екатеринбург ул.Азина 27");
            tableText(table, 10, 0, 10, false, "Грузополучатель, контактное лицо");
            tableText(table, 10, 1, 10, false, "-");
            tableText(table, 11, 0, 10, true, "2. Дата и время подачи под выгрузку");
            tableText(table, 11, 1, 10, false, "-");
            tableText(table, 12, 0, 10, false, "Адрес второго места выгрузки");
            tableText(table, 12, 1, 10, false, "-");
            tableText(table, 13, 0, 10, true, "Грузополучатель, контактное лицо");
            tableText(table, 13, 1, 10, false, "-");
            tableText(table, 14, 0, 10, true, "Стоимость перевозки");
            tableText(table, 14, 1, 10, true, "Без НДС 79000руб");
            tableText(table, 15, 0, 10, true, "Сроки оплаты и форма");
            tableText(table, 15, 1, 10, true, "по оригиналам ТТН и бухгалтерских документов 10-15 б.д.");
            tableText(table, 16, 0, 10, true, "Данные а/м и п/п");
            tableText(table, 16, 1, 10, false, "DAF У927ОМ116. п/п АУ1404/16");
            tableText(table, 17, 0, 10, true, "Ф.И.О. водителя");
            tableText(table, 17, 1, 10, false, "Ахкиямов Ильнар Минемухаметови");
            tableText(table, 18, 0, 10, true, "Паспорт водителя");
            tableText(table, 18, 1, 8, false, "Паспорт 9203 903001 ОВД МЕНДЕЛЕЕВСКОГО РАЙОНА РЕСПУБЛИКИ ТАТАРСТАН 20.12.2002");
            tableText(table, 19, 0, 10, true, "К. тел. Водителя");
            tableText(table, 19, 1, 10, false, "89274508655");
            createTextWithoutSpacing(docxModel, ParagraphAlignment.LEFT, "Данная перевозка осуществляется в соответствии с условиями конвенций КДПГ, МДП (при международных перевозках), Гражданского Кодекса РФ, Устава Автомобильного Транспорта и Договора-заявки.", 8, true);
            createTextWithoutSpacing(docxModel, ParagraphAlignment.LEFT, "Данная Заявка при отсутствии долгосрочного Договора между Заказчиком и Исполнителем, имеет силу Договора на разовую перевозку. Факсимильная копия данной Заявки имеет юридическую силу.", 8, false);
            createTextWithoutSpacing(docxModel, ParagraphAlignment.LEFT, "На погрузку/выгрузку  выделяется  24 часа с момента прибытия а/м.", 8, false);
            createTextWithoutSpacing(docxModel, ParagraphAlignment.LEFT, "Заказчик обязан:  ", 8, false, false, "1) Согласовывать информацию о грузе с Исполнителем.", 8, true, false);
            createTextWithoutSpacing(docxModel, ParagraphAlignment.LEFT, "                                2) Заказчик вправе отказаться от перевозки за 12 часов до погрузки без финансовых последствий для себя.", 8, true);
            createTextWithoutSpacing(docxModel, ParagraphAlignment.LEFT, "                                3) При оплате за перевозку на банковскую карту – комиссия банка (%) взымается за счет получателя.", 8, true);
            createTextWithoutSpacing(docxModel, ParagraphAlignment.LEFT, "                                4) Простой а/м по вине Заказчика оплачивается 1000 руб/сут. при условии документального подтверждения.", 8, true);
            createTextWithoutSpacing(docxModel, ParagraphAlignment.LEFT, "Исполнитель обязан:", 8, true);
            createTextWithoutSpacing(docxModel, ParagraphAlignment.LEFT, "1)  Подавать под погрузку технически исправные автомашины. Заказчик вправе отказаться от автомашин, непригодных для перевозки соответствующего груза.", 8, false);
            createTextWithoutSpacing(docxModel, ParagraphAlignment.LEFT, "2) В случае возникновения неисправности в транспортном средстве, во время оказания услуг или схода транспортного средства, уведомить Заказчика и заменить за свой счет неисправное транспортное средство равноценным исправным без изменения установленных сроков оказания услуг.", 8, false);
            createTextWithoutSpacing(docxModel, ParagraphAlignment.LEFT, "3) В случае отказа Исполнителя от загрузки по подтвержденной заявке, Исполнитель выплачивает Заказчику штраф в размере 20% от стоимости перевозки.", 8, false);
            createTextWithoutSpacing(docxModel, ParagraphAlignment.LEFT, "Водитель Исполнителя обязан:", 8, true);
            createTextWithoutSpacing(docxModel, ParagraphAlignment.LEFT, "1) Следить за погрузкой/выгрузкой, а также за креплением груза. Принимать/сдавать груз по весу, количеству мест, целостности упаковки, наименованию и номенклатуре, согласно накладной.", 8, false);
            createTextWithoutSpacing(docxModel, ParagraphAlignment.LEFT, "2) Сообщить Заказчику о нарушениях в правилах укладки груза, угрожающих его сохранности до начала перевозки.", 8, false);
            createTextWithoutSpacing(docxModel, ParagraphAlignment.LEFT, "3) Сообщить Заказчику о перегрузе, если перегруз не был ранее обговорен в Заявке.", 8, false);
            createTextWithoutSpacing(docxModel, ParagraphAlignment.LEFT, "4) Известить Заказчика, прежде чем подписать Акт или Протокол порчи, несоответствия груза.", 8, false);
            createTextWithoutSpacing(docxModel, ParagraphAlignment.LEFT, "5) Водитель несет полную материальную ответственность за целостность и сохранность вверенного ему груза.", 8, false);
            createTextWithoutSpacing(docxModel, ParagraphAlignment.LEFT, "6) Оригиналы документов выслать по адресу: ", 8, false, false, "445054, Самарская область, г. Тольятти, ул. Мира, д.93, кв.12", 8, true, false);
            createText(docxModel, ParagraphAlignment.LEFT, "Дополнительные условия: ", 8, true, false, "Место разгрузки строго по ТТН/ТН. В случае разгрузки не по адресу, ИП Фокин.А.Б. оплачивает ООО «Союз Транс» ", 8, false, false, "100% стоимость груза !!!", 8, true, false);
            createText(docxModel, ParagraphAlignment.LEFT, "", 11, false);
            createText(docxModel, ParagraphAlignment.LEFT, "", 11, false);
            createTextWithoutSpacing(docxModel, ParagraphAlignment.CENTER, "Исполнитель:_____________________                                     Заказчик:_____________________", 11, false);
            createText(docxModel, ParagraphAlignment.LEFT, "                                                                                  м. п.                                                                                                                                  м. п", 8, false);
            FileOutputStream outputStream = new FileOutputStream("Word Test.docx");
            docxModel.write(outputStream);
            outputStream.close();
        } catch (Exception var7) {
            var7.printStackTrace();
        }
    }
    private  CTP createFooterModel(String footerContent) {
        CTP ctpFooterModel = CTP.Factory.newInstance();
        CTR ctrFooterModel = ctpFooterModel.addNewR();
        CTText cttFooter = ctrFooterModel.addNewT();
        cttFooter.setStringValue(footerContent);
        return ctpFooterModel;
    }

    private  CTP createHeaderModel(String headerContent) {
        CTP ctpHeaderModel = CTP.Factory.newInstance();
        CTR ctrHeaderModel = ctpHeaderModel.addNewR();
        CTText cttHeader = ctrHeaderModel.addNewT();
        cttHeader.setStringValue(headerContent);
        return ctpHeaderModel;
    }

    private  void createText(XWPFDocument docxModel, ParagraphAlignment alignment, String text, int size, boolean bold) {
        XWPFParagraph bodyParagraph = docxModel.createParagraph();
        bodyParagraph.setAlignment(alignment);
        XWPFRun paragraphConfig = bodyParagraph.createRun();
        paragraphConfig.setFontSize(size);
        paragraphConfig.setColor("000000");
        paragraphConfig.setFontFamily("Times New Roman");
        paragraphConfig.setText(text);
        paragraphConfig.setBold(bold);
    }

    private  void createText(XWPFDocument docxModel, ParagraphAlignment position, String text, int size, boolean bold, boolean line, String sText, int sSize, boolean sBold, boolean sLine) {
        XWPFParagraph bodyParagraph = docxModel.createParagraph();
        bodyParagraph.setAlignment(position);
        XWPFRun paragraphConfig = bodyParagraph.createRun();
        paragraphConfig.setFontSize(size);
        if (line) {
            paragraphConfig.setUnderline(UnderlinePatterns.SINGLE);
        }

        paragraphConfig.setColor("000000");
        paragraphConfig.setFontFamily("Times New Roman");
        paragraphConfig.setBold(bold);
        paragraphConfig.setText(text);
        XWPFRun paragraphConfig2 = bodyParagraph.createRun();
        paragraphConfig2.setFontSize(sSize);
        paragraphConfig2.setColor("000000");
        paragraphConfig2.setFontFamily("Times New Roman");
        paragraphConfig2.setBold(sBold);
        if (sLine) {
            paragraphConfig2.setUnderline(UnderlinePatterns.SINGLE);
        }

        paragraphConfig2.setText(sText);
    }

    private  void createText(XWPFDocument docxModel, ParagraphAlignment position, String text, int size, boolean bold, boolean line, String sText, int sSize, boolean sBold, boolean sLine, String tText, int tSize, boolean tBold, boolean tLine) {
        XWPFParagraph bodyParagraph = docxModel.createParagraph();
        bodyParagraph.setAlignment(position);
        XWPFRun paragraphConfig = bodyParagraph.createRun();
        paragraphConfig.setFontSize(size);
        if (line) {
            paragraphConfig.setUnderline(UnderlinePatterns.SINGLE);
        }

        paragraphConfig.setColor("000000");
        paragraphConfig.setFontFamily("Times New Roman");
        paragraphConfig.setBold(bold);
        paragraphConfig.setText(text);
        XWPFRun paragraphConfig2 = bodyParagraph.createRun();
        paragraphConfig2.setFontSize(sSize);
        paragraphConfig2.setColor("000000");
        paragraphConfig2.setFontFamily("Times New Roman");
        paragraphConfig2.setBold(sBold);
        if (sLine) {
            paragraphConfig2.setUnderline(UnderlinePatterns.SINGLE);
        }

        paragraphConfig2.setText(sText);
        XWPFRun paragraphConfig3 = bodyParagraph.createRun();
        paragraphConfig3.setFontSize(tSize);
        paragraphConfig3.setColor("000000");
        paragraphConfig3.setFontFamily("Times New Roman");
        paragraphConfig3.setBold(tBold);
        if (tLine) {
            paragraphConfig2.setUnderline(UnderlinePatterns.SINGLE);
        }

        paragraphConfig3.setText(tText);
    }

    private  void createTextWithoutSpacing(XWPFDocument docxModel, ParagraphAlignment alignment, String text, int size, boolean bold) {
        XWPFParagraph bodyParagraph = docxModel.createParagraph();
        bodyParagraph.setAlignment(alignment);
        XWPFRun paragraphConfig = bodyParagraph.createRun();
        paragraphConfig.setFontSize(size);
        paragraphConfig.setColor("000000");
        paragraphConfig.setFontFamily("Times New Roman");
        paragraphConfig.setText(text);
        paragraphConfig.setBold(bold);
        bodyParagraph.setSpacingAfter(0);
    }
    private  void createTextWithoutSpacing(XWPFDocument docxModel, ParagraphAlignment position, String text, int size, boolean bold, boolean line, String sText, int sSize, boolean sBold, boolean sLine) {
        XWPFParagraph bodyParagraph = docxModel.createParagraph();
        bodyParagraph.setAlignment(position);
        XWPFRun paragraphConfig = bodyParagraph.createRun();
        paragraphConfig.setFontSize(size);
        if (line) {
            paragraphConfig.setUnderline(UnderlinePatterns.SINGLE);
        }

        paragraphConfig.setColor("000000");
        paragraphConfig.setFontFamily("Times New Roman");
        paragraphConfig.setBold(bold);
        paragraphConfig.setText(text);
        XWPFRun paragraphConfig2 = bodyParagraph.createRun();
        paragraphConfig2.setFontSize(sSize);
        paragraphConfig2.setColor("000000");
        paragraphConfig2.setFontFamily("Times New Roman");
        paragraphConfig2.setBold(sBold);
        bodyParagraph.setSpacingAfter(0);
        if (sLine) {
            paragraphConfig2.setUnderline(UnderlinePatterns.SINGLE);
        }

        paragraphConfig2.setText(sText);
    }



    private  void createText(XWPFDocument docxModel, ParagraphAlignment position, String text, int size, boolean bold, boolean line, String sText, int sSize, boolean sBold, boolean sLine, String tText, int tSize, boolean tBold, boolean tLine, String fText, int fSize, boolean fBold, boolean fLine) {
        XWPFParagraph bodyParagraph = docxModel.createParagraph();
        bodyParagraph.setAlignment(position);
        XWPFRun paragraphConfig = bodyParagraph.createRun();
        paragraphConfig.setFontSize(size);
        if (line) {
            paragraphConfig.setUnderline(UnderlinePatterns.SINGLE);
        }

        paragraphConfig.setColor("000000");
        paragraphConfig.setFontFamily("Times New Roman");
        paragraphConfig.setBold(bold);
        paragraphConfig.setText(text);
        XWPFRun paragraphConfig2 = bodyParagraph.createRun();
        paragraphConfig2.setFontSize(sSize);
        paragraphConfig2.setColor("000000");
        paragraphConfig2.setFontFamily("Times New Roman");
        paragraphConfig2.setBold(sBold);
        if (sLine) {
            paragraphConfig2.setUnderline(UnderlinePatterns.SINGLE);
        }

        paragraphConfig2.setText(sText);
        XWPFRun paragraphConfig3 = bodyParagraph.createRun();
        paragraphConfig3.setFontSize(tSize);
        paragraphConfig3.setColor("000000");
        paragraphConfig3.setFontFamily("Times New Roman");
        paragraphConfig3.setBold(tBold);
        if (tLine) {
            paragraphConfig2.setUnderline(UnderlinePatterns.SINGLE);
        }

        paragraphConfig3.setText(tText);
    }

    private  void tableText(XWPFTable table, int row, int cell, int size, boolean bold, String text) {
        XWPFParagraph paragraph = table.getRow(row).getCell(cell).addParagraph();
        paragraph.setSpacingAfter(10);
        XWPFRun paragraphConfig = paragraph.createRun();
        paragraphConfig.setFontSize(size);
        paragraphConfig.setColor("000000");
        paragraphConfig.setFontFamily("Times New Roman");
        paragraphConfig.setBold(bold);
        paragraphConfig.setText(text);
        table.getRow(row).getCell(cell).removeParagraph(0);
        table.getRow(row).getCell(cell).setParagraph(paragraph);
    }
}

package com.company;

import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDPage;
import org.apache.pdfbox.pdmodel.PDPageContentStream;
import org.apache.pdfbox.pdmodel.font.PDFont;
import org.apache.pdfbox.pdmodel.font.PDType1Font;
import org.apache.pdfbox.text.PDFTextStripper;
import org.apache.pdfbox.text.PDFTextStripperByArea;

import java.io.File;
import java.io.IOException;

public class PdfFile {
    public PdfFile(String directory) {
        try {
            readPdf(directory);
        } catch (IOException e) {
            e.printStackTrace();
        }

    }

    private void readPdf(String directory) throws IOException {
        File file = new File("C:\\Users\\vladl\\IdeaProjects\\Template Manager\\pdf.pdf");
        PDDocument document = PDDocument.load(file);
        System.out.println("Load");

        /*
        //Adding a blank page to the document
        document.addPage(new PDPage());
        */
        PDPage pdPage = new  PDPage();
        PDFTextStripper textStripper = new PDFTextStripper();
        String text = textStripper.getText(document);
        System.out.println(text);


        //Saving the document
        document.save("sample.pdf");

        //Closing the document
        document.close();
    }


}
    


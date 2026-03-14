package com.example;

import java.awt.image.BufferedImage;
import java.io.File;
import java.io.IOException;

import org.apache.pdfbox.Loader;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDPage;
import org.apache.pdfbox.pdmodel.PDPageContentStream;
import org.apache.pdfbox.pdmodel.common.PDRectangle;
import org.apache.pdfbox.pdmodel.graphics.image.PDImageXObject;
import org.apache.pdfbox.rendering.ImageType;
import org.apache.pdfbox.rendering.PDFRenderer;

public class PdfTextToPdfImg {

	public static void convertPdfToImagePdf(String sourcePath, String destinationPath) throws IOException {
        try (PDDocument sourceDocument =  Loader.loadPDF(new File(sourcePath));
             PDDocument destinationDocument = new PDDocument()) {

            PDFRenderer pdfRenderer = new PDFRenderer(sourceDocument);

            System.out.println("Source: " + sourcePath);
            System.out.println("NumberOfPages: " + sourceDocument.getNumberOfPages());

            for (int pageNum = 0; pageNum < sourceDocument.getNumberOfPages(); ++pageNum) {
                
                BufferedImage bufferedImage = pdfRenderer.renderImageWithDPI(pageNum, 300, ImageType.RGB);


                PDImageXObject pdImage = org.apache.pdfbox.pdmodel.graphics.image.LosslessFactory.createFromImage(destinationDocument, bufferedImage);

                PDRectangle newPageSize = new PDRectangle(pdImage.getWidth(), pdImage.getHeight());
                PDPage newPage = new PDPage(newPageSize);
                destinationDocument.addPage(newPage);

                try (PDPageContentStream contentStream = new PDPageContentStream(destinationDocument, newPage)) {
                    contentStream.drawImage(pdImage, 0, 0, pdImage.getWidth(), pdImage.getHeight());
                }

                System.out.println("Page " + (pageNum + 1) + " Done");
            }

            destinationDocument.save(destinationPath);
            System.out.println("Output Path: " + destinationPath);
        }
    }

    public static void main(String[] args) {
        String inputPdfPath = "D:/Working/PDF/rcp25022025awn_co_SEAL_1.pdf";
        String outputPdfPath = "D:/Working/PDF/output.pdf";
        // -------------------------

        File inputFile = new File(inputPdfPath);
        if (!inputFile.exists()) {
            System.out.println("Input file not found: " + inputPdfPath);
            return;
        }

        try {
            convertPdfToImagePdf(inputPdfPath, outputPdfPath);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

}

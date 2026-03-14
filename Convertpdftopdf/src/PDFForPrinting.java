import com.itextpdf.kernel.pdf.PdfDocument;
import com.itextpdf.kernel.pdf.PdfReader;
import com.itextpdf.kernel.pdf.PdfWriter;
import com.itextpdf.kernel.pdf.canvas.PdfCanvas;

public class PDFForPrinting {
    public static void main(String[] args) {
        try {
            // Load the existing PDF
            PdfReader reader = new PdfReader("input.pdf");
            PdfWriter writer = new PdfWriter("print_ready.pdf");
            PdfDocument pdfDoc = new PdfDocument(reader, writer);

            // Add crop marks or other print-specific modifications
            PdfCanvas canvas = new PdfCanvas(pdfDoc.getFirstPage());
            canvas.rectangle(36, 36, 540, 756); // Example crop box
            canvas.stroke();

            // Close the document
            pdfDoc.close();
            System.out.println("PDF prepared for printing.");
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}

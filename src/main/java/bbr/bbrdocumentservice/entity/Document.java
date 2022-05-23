package bbr.bbrdocumentservice.entity;

import com.aspose.cells.SaveFormat;
import com.aspose.cells.Workbook;
import com.itextpdf.text.BaseColor;
import com.itextpdf.text.Rectangle;
import com.itextpdf.text.pdf.PdfReader;
import com.itextpdf.text.pdf.PdfStamper;
import com.itextpdf.text.pdf.pdfcleanup.PdfCleanUpLocation;
import com.itextpdf.text.pdf.pdfcleanup.PdfCleanUpProcessor;

import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.List;

public class Document {

    /**
     * Example:
     * src = "Book1.xlsx"
     * dst = "Excel-to-PDF.pdf"
     */
    public static void convertXlsxToPdf(String src, String dst) {

        try {
            Workbook workbook = new Workbook(src);
            workbook.save(dst, SaveFormat.PDF);
        } catch (Exception e) {
            e.printStackTrace();
        }

    }

    /**
     * @param src "input.pdf"
     * @param dst "output.pdf"
     */
    public static void removeAsposeText(String src, String dst) {
        try {
            PdfReader reader = new PdfReader(src);
            PdfStamper stamper = new PdfStamper(reader, new FileOutputStream(dst));
            List<PdfCleanUpLocation> cleanUpLocations = new ArrayList<PdfCleanUpLocation>();
            cleanUpLocations.add(new PdfCleanUpLocation(1, new Rectangle(1f, 820f, 593f, 841f), BaseColor.WHITE));
            PdfCleanUpProcessor cleaner = new PdfCleanUpProcessor(cleanUpLocations, stamper);
            cleaner.cleanUp();
            stamper.close();
            reader.close();
        } catch (Exception e) {
            e.printStackTrace();
        }

    }
}

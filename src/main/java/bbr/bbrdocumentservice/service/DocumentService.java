package bbr.bbrdocumentservice.service;


import bbr.bbrdocumentservice.entity.Document;
import bbr.bbrdocumentservice.entity.InvoiceDTO;
import bbr.bbrdocumentservice.entity.ItemDTO;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.List;

@Service
public class DocumentService {

    private static final String RON_INVOICE_NAME = ".\\invoices\\ron.xlsx";
    private static final String EUR_INVOICE_NAME = ".\\invoices\\eur.xlsx";
    private static final String EXCELS_LOCATION = ".\\invoices\\excels\\";
    private static final String ASPOSES_LOCATION = ".\\invoices\\asposes\\";
    private static final String PDFS_LOCATION = ".\\invoices\\pdfs\\";
    private static final Integer NUMBER_OF_ITEMS = 12;


    public File createPdf(InvoiceDTO invoiceDTO) {

        createXlsx(invoiceDTO);
        Document.convertXlsxToPdf(generateXlxsName(invoiceDTO), generateAsposeName(invoiceDTO));
        Document.removeAsposeText(generateAsposeName(invoiceDTO), generatePdfName(invoiceDTO));

        File file = new File(generatePdfName(invoiceDTO));
        return file;

    }

    private String generateXlxsName(InvoiceDTO invoiceDTO) {
        return EXCELS_LOCATION + invoiceDTO.getInvoiceCode() + ".xlsx";
    }

    private String generateAsposeName(InvoiceDTO invoiceDTO) {
        return ASPOSES_LOCATION + invoiceDTO.getInvoiceCode() + ".pdf";
    }

    private String generatePdfName(InvoiceDTO invoiceDTO) {
        return PDFS_LOCATION + invoiceDTO.getInvoiceCode() + ".pdf";
    }

    public void createXlsx(InvoiceDTO invoiceDTO) {
        try {


            /**
             * A0  B1  C2  D3  E4  F5  G6  H7  I8  J9
             */
            if (invoiceDTO.getIsEur()) {

                FileInputStream file = new FileInputStream(new File(EUR_INVOICE_NAME));
                Workbook workbook = new XSSFWorkbook(file);

                if (invoiceDTO.getIsStorno()) {
                    workbook.getSheetAt(0).getRow(0).getCell(4).setCellValue("STORNO");
                    workbook.getSheetAt(0).getRow(0).getCell(6).setCellValue(invoiceDTO.getStornoCode());
                } else {
                    workbook.getSheetAt(0).getRow(0).getCell(4).setCellValue("Factura");
                    workbook.getSheetAt(0).getRow(0).getCell(6).setCellValue("");
                }

                workbook.getSheetAt(0).getRow(0).getCell(8).setCellValue(invoiceDTO.getInvoiceCode());
                workbook.getSheetAt(0).getRow(1).getCell(5).setCellValue(invoiceDTO.getCreationDate());
                workbook.getSheetAt(0).getRow(1).getCell(8).setCellValue("Cota TVA " + invoiceDTO.getTva() + "%");
                workbook.getSheetAt(0).getRow(3).getCell(8).setCellValue(invoiceDTO.getPriceTotal());
                workbook.getSheetAt(0).getRow(4).getCell(8).setCellValue(invoiceDTO.getPriceTotal() * invoiceDTO.getEuro());

                workbook.getSheetAt(0).getRow(8).getCell(4).setCellValue(invoiceDTO.getCompany().getName());
                workbook.getSheetAt(0).getRow(10).getCell(5).setCellValue(invoiceDTO.getCompany().getCif());
                workbook.getSheetAt(0).getRow(11).getCell(5).setCellValue(invoiceDTO.getCompany().getReg());
                workbook.getSheetAt(0).getRow(12).getCell(5).setCellValue(invoiceDTO.getCompany().getAddress());
                workbook.getSheetAt(0).getRow(13).getCell(5).setCellValue(invoiceDTO.getCompany().getEmail());

                List<ItemDTO> itemDTOList = invoiceDTO.getItems();

                for (int i = 0; i < NUMBER_OF_ITEMS; i++) {

                    if (i < itemDTOList.size()) {
                        ItemDTO item = itemDTOList.get(i);
                        workbook.getSheetAt(0).getRow(22 + i).getCell(0).setCellValue(i + 1);
                        workbook.getSheetAt(0).getRow(22 + i).getCell(1).setCellValue(item.getName());
                        workbook.getSheetAt(0).getRow(22 + i).getCell(3).setCellValue(0);
                        workbook.getSheetAt(0).getRow(22 + i).getCell(4).setCellValue(item.getQuantity());
                        workbook.getSheetAt(0).getRow(22 + i).getCell(5).setCellValue(item.getPrice());
                        workbook.getSheetAt(0).getRow(22 + i).getCell(6).setCellValue(item.getQuantity() * item.getPrice());
                        workbook.getSheetAt(0).getRow(22 + i).getCell(8).setCellValue(invoiceDTO.getTva() / 100 * (item.getQuantity() * item.getPrice()));
                    } else {
                        workbook.getSheetAt(0).getRow(22 + i).getCell(0).setCellValue("");
                        workbook.getSheetAt(0).getRow(22 + i).getCell(1).setCellValue("");
                        workbook.getSheetAt(0).getRow(22 + i).getCell(3).setCellValue("");
                        workbook.getSheetAt(0).getRow(22 + i).getCell(4).setCellValue("");
                        workbook.getSheetAt(0).getRow(22 + i).getCell(5).setCellValue("");
                        workbook.getSheetAt(0).getRow(22 + i).getCell(6).setCellValue("");
                        workbook.getSheetAt(0).getRow(22 + i).getCell(8).setCellValue("");
                    }
                }

                workbook.getSheetAt(0).getRow(36).getCell(2).setCellValue(invoiceDTO.getEuro());
                workbook.getSheetAt(0).getRow(36).getCell(6).setCellValue(invoiceDTO.getPriceWithoutTva());
                workbook.getSheetAt(0).getRow(36).getCell(8).setCellValue(invoiceDTO.getPriceTva());
                workbook.getSheetAt(0).getRow(37).getCell(6).setCellValue(invoiceDTO.getPriceWithoutTva() * invoiceDTO.getEuro());
                workbook.getSheetAt(0).getRow(37).getCell(8).setCellValue(invoiceDTO.getPriceTva() * invoiceDTO.getEuro());

                workbook.getSheetAt(0).getRow(44).getCell(8).setCellValue(invoiceDTO.getPriceTotal());
                workbook.getSheetAt(0).getRow(45).getCell(8).setCellValue(invoiceDTO.getPriceTotal() * invoiceDTO.getEuro());

                workbook.getSheetAt(0).getRow(47).getCell(8).setCellValue(invoiceDTO.getDueDate());

                FileOutputStream outputStream = new FileOutputStream(new File(generateXlxsName(invoiceDTO)));
                workbook.write(outputStream);
                workbook.close();

                /**
                 * A0  B1  C2  D3  E4  F5  G6  H7  I8  J9
                 */
            } else {

                FileInputStream file = new FileInputStream(new File(RON_INVOICE_NAME));
                Workbook workbook = new XSSFWorkbook(file);

                if (invoiceDTO.getIsStorno()) {
                    workbook.getSheetAt(0).getRow(0).getCell(4).setCellValue("STORNO");
                    workbook.getSheetAt(0).getRow(0).getCell(6).setCellValue(invoiceDTO.getStornoCode());
                } else {
                    workbook.getSheetAt(0).getRow(0).getCell(4).setCellValue("Factura");
                    workbook.getSheetAt(0).getRow(0).getCell(6).setCellValue("");
                }

                workbook.getSheetAt(0).getRow(0).getCell(8).setCellValue(invoiceDTO.getInvoiceCode());
                workbook.getSheetAt(0).getRow(1).getCell(5).setCellValue(invoiceDTO.getCreationDate());
                workbook.getSheetAt(0).getRow(1).getCell(8).setCellValue("Cota TVA " + invoiceDTO.getTva() + "%");
                workbook.getSheetAt(0).getRow(3).getCell(8).setCellValue(invoiceDTO.getPriceTotal());

                workbook.getSheetAt(0).getRow(8).getCell(4).setCellValue(invoiceDTO.getCompany().getName());
                workbook.getSheetAt(0).getRow(10).getCell(5).setCellValue(invoiceDTO.getCompany().getCif());
                workbook.getSheetAt(0).getRow(11).getCell(5).setCellValue(invoiceDTO.getCompany().getReg());
                workbook.getSheetAt(0).getRow(12).getCell(5).setCellValue(invoiceDTO.getCompany().getAddress());
                workbook.getSheetAt(0).getRow(13).getCell(5).setCellValue(invoiceDTO.getCompany().getEmail());

                List<ItemDTO> itemDTOList = invoiceDTO.getItems();

                for (int i = 0; i < NUMBER_OF_ITEMS; i++) {

                    if (i < itemDTOList.size()) {
                        ItemDTO item = itemDTOList.get(i);
                        workbook.getSheetAt(0).getRow(22 + i).getCell(0).setCellValue(i + 1);
                        workbook.getSheetAt(0).getRow(22 + i).getCell(1).setCellValue(item.getName());
                        workbook.getSheetAt(0).getRow(22 + i).getCell(3).setCellValue(0);
                        workbook.getSheetAt(0).getRow(22 + i).getCell(4).setCellValue(item.getQuantity());
                        workbook.getSheetAt(0).getRow(22 + i).getCell(5).setCellValue(item.getPrice());
                        workbook.getSheetAt(0).getRow(22 + i).getCell(6).setCellValue(item.getQuantity() * item.getPrice());
                        workbook.getSheetAt(0).getRow(22 + i).getCell(8).setCellValue(invoiceDTO.getTva() / 100 * (item.getQuantity() * item.getPrice()));
                    } else {
                        workbook.getSheetAt(0).getRow(22 + i).getCell(0).setCellValue("");
                        workbook.getSheetAt(0).getRow(22 + i).getCell(1).setCellValue("");
                        workbook.getSheetAt(0).getRow(22 + i).getCell(3).setCellValue("");
                        workbook.getSheetAt(0).getRow(22 + i).getCell(4).setCellValue("");
                        workbook.getSheetAt(0).getRow(22 + i).getCell(5).setCellValue("");
                        workbook.getSheetAt(0).getRow(22 + i).getCell(6).setCellValue("");
                        workbook.getSheetAt(0).getRow(22 + i).getCell(8).setCellValue("");
                    }
                }

                workbook.getSheetAt(0).getRow(36).getCell(6).setCellValue(invoiceDTO.getPriceWithoutTva());
                workbook.getSheetAt(0).getRow(36).getCell(8).setCellValue(invoiceDTO.getPriceTva());

                workbook.getSheetAt(0).getRow(44).getCell(8).setCellValue(invoiceDTO.getPriceTotal());

                workbook.getSheetAt(0).getRow(47).getCell(8).setCellValue(invoiceDTO.getDueDate());

                FileOutputStream outputStream = new FileOutputStream(new File(generateXlxsName(invoiceDTO)));
                workbook.write(outputStream);
                workbook.close();
            }


        } catch (Exception e) {
            e.printStackTrace();
        }
    }


}

package bbr.bbrdocumentservice.controller;


import bbr.bbrdocumentservice.entity.InvoiceDTO;
import bbr.bbrdocumentservice.service.DocumentService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpStatus;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;

import java.io.File;
import java.io.IOException;
import java.nio.file.Files;

@RestController
public class DocumentController {

    @Autowired
    private DocumentService documentService;

    @GetMapping(path = "/welcome")
    public String home() {
        return "Welcome to [ Document Service ] !";
    }


    @PostMapping(path = "/get-pdf")
    public ResponseEntity<byte[]> getPDF(@RequestBody InvoiceDTO invoiceDTO) {

        // retrieve contents of "C:/tmp/report.pdf" that were written in showHelp
        File file = documentService.createPdf(invoiceDTO);
        byte[] contents = new byte[0];
        try {
            contents = Files.readAllBytes(file.toPath());
        } catch (IOException e) {
            e.printStackTrace();
        }

        HttpHeaders headers = new HttpHeaders();
        headers.setContentType(MediaType.APPLICATION_PDF);

        // Here you have to set the actual filename of your pdf
        String filename = "output.pdf";
        headers.setContentDispositionFormData(filename, filename);
        headers.setCacheControl("must-revalidate, post-check=0, pre-check=0");

        ResponseEntity<byte[]> response = new ResponseEntity<>(contents, headers, HttpStatus.OK);

        return response;
    }

}

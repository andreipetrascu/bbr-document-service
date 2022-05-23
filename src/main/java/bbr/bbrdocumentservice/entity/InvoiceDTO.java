package bbr.bbrdocumentservice.entity;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;
import lombok.ToString;

import java.time.LocalDate;
import java.util.List;

@Data
@AllArgsConstructor
@NoArgsConstructor
@ToString
public class InvoiceDTO {

    private String invoiceCode;

    private LocalDate creationDate;

    private LocalDate dueDate;

    private Float euro;

    private Float tva;

    private Float priceWithoutTva;

    private Float priceTva;

    private Float priceTotal;

    private Boolean isEur;

    private Boolean isStorno;

    private CompanyDTO company;

    private List<ItemDTO> items;

    private String stornoCode;

}

package bbr.bbrdocumentservice.entity;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;
import lombok.ToString;

@Data
@AllArgsConstructor
@NoArgsConstructor
@ToString
public class CompanyDTO {
    private String name;

    private String cif;

    private String reg;

    private String email;

    private String address;

}

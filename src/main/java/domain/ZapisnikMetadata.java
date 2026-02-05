package domain;

import java.time.LocalDate;
import java.time.format.DateTimeFormatter;

//Zaglavlje samog fajla, podaci koji nisu vezani za stavke

public class ZapisnikMetadata {
    private final LocalDate datum;
    private final LocalDate datumSlash;
    private final String kBroj;
    private final String broj;
    private final String nazivRaskrsnice;
    
    public ZapisnikMetadata(
            LocalDate datum,
            LocalDate datumSlash,
            String kBroj,
            String broj,
            String nazivRaskrsnice
    ) {
        this.datum = datum;
        this.datumSlash = datumSlash;
        this.kBroj = kBroj;
        this.broj = broj;
        this.nazivRaskrsnice = nazivRaskrsnice;
    }
    
    public LocalDate getDatum() {
        return datum;
    }
    
    public String getDatumFormatiranDot() {
        return datum.format(DateTimeFormatter.ofPattern("dd.MM.yyyy."));
    }
    
    public String getDatumFormatiranSlash(){
        return datumSlash.format(DateTimeFormatter.ofPattern("dd/MM/yyyy."));
    }
    
    public String getKBroj() {
        return kBroj;
    }
    
    public String getBroj() {
        return broj;
    }
    
    public String getNazivRaskrsnice() {
        return nazivRaskrsnice;
    }
}

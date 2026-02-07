package service;

import domain.Stavka;
import excel.CenovnikLookup;
import excel.CenovnikLookup.CenovnikItem;
import repository.RaskrsniceRepository;

import java.math.BigDecimal;

public class ExcelReaderService {
    
    private final CenovnikLookup cenovnikLookup;
    private final RaskrsniceRepository raskrsniceRepository;
    
    public ExcelReaderService(CenovnikLookup cenovnikLookup, RaskrsniceRepository raskrsniceRepository) {
        this.cenovnikLookup = cenovnikLookup;
        this.raskrsniceRepository = raskrsniceRepository;
    }
    
    public String getNazivRaskrsnice(String kBroj) {
        if (kBroj == null || kBroj.isBlank()) {
            throw new IllegalArgumentException("K broj ne sme biti prazan.");
        }
        return raskrsniceRepository.getNazivRaskrsnice(kBroj.trim());
    }
    
    public Stavka napraviStavku(int redniBroj, BigDecimal kolicina) {
        if (redniBroj <= 0) {
            throw new IllegalArgumentException("Redni broj stavke mora biti > 0.");
        }
        if (kolicina == null || kolicina.compareTo(BigDecimal.ZERO) <= 0) {
            throw new IllegalArgumentException("Količina mora biti veća od 0.");
        }
        
        CenovnikItem item = cenovnikLookup.findByRedniBroj(redniBroj)
                .orElseThrow(() -> new IllegalArgumentException(
                        "Ne postoji stavka sa RB = " + redniBroj + " u cenovniku."
                ));
        
        return new Stavka(
                redniBroj,
                item.getNaziv(),
                item.getJm(),
                kolicina,
                item.getCena()
        );
    }
}

package excel;

import java.math.BigDecimal;
import java.util.Optional;

//Apstrakcija cenovnika

public interface CenovnikLookup {
    Optional<CenovnikItem> findByRedniBroj(int redniBroj);
    
    final class CenovnikItem {
        private final String naziv;
        private final String jm;
        private final BigDecimal cena;
        
        public CenovnikItem(String naziv,String jedinicaMere, BigDecimal cena) {
            this.naziv = naziv;
            this.jm = jedinicaMere;
            this.cena = cena;
        }
        
        public String getNaziv() {
            return naziv;
        }
        
        public BigDecimal getCena() {
            return cena;
        }
        
        public String getJm() {
            return jm;
        }
    }
}

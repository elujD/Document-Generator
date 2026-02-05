import java.math.BigDecimal;

public class Stavka {
    
    public final int rb;
    public final String modulUsluga;
    public final String jedinicaMere;
    public final BigDecimal kolicina;
    public final BigDecimal cena;
    public final BigDecimal ukupno;
    
    public Stavka(int rb,
                  String modulUsluga,
                  String jedinicaMere,
                  BigDecimal kolicina,
                  BigDecimal cena) {
        
        this.rb = rb;
        this.modulUsluga = modulUsluga;
        this.jedinicaMere = jedinicaMere;
        this.kolicina = kolicina;
        this.cena = cena;
        this.ukupno = kolicina.multiply(cena);
    }
}

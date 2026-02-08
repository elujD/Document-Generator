package service;

import domain.ZapisnikMetadata;
import domain.Stavka;
import word.WordTabelaPopunjavanje;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import java.math.BigDecimal;
import java.util.*;

public class ZapisnikService {
    
    private final ExcelReaderService excelReaderService;
    
    public ZapisnikService(ExcelReaderService excelReaderService) {
        this.excelReaderService = Objects.requireNonNull(excelReaderService, "excelReaderService ne sme biti null");
    }
    
    public void generisiZapisnik(XWPFDocument doc,
                                 ZapisnikMetadata metadata,
                                 List<Integer> redniBrojevi,
                                 List<BigDecimal> kolicine) {
        
        if (doc == null) throw new IllegalArgumentException("Word dokument (doc) ne sme biti null.");
        if (metadata == null) throw new IllegalArgumentException("Metadata ne sme biti null.");
        if (redniBrojevi == null || kolicine == null) throw new IllegalArgumentException("Liste ne smeju biti null.");
        if (redniBrojevi.size() != kolicine.size()) throw new IllegalArgumentException("RB i količine moraju imati isti broj elemenata.");
        if (redniBrojevi.isEmpty()) throw new IllegalArgumentException("Moraš uneti bar jednu stavku.");
        
        List<Stavka> stavke = new ArrayList<>(redniBrojevi.size());
        for (int i = 0; i < redniBrojevi.size(); i++) {
            int rb = redniBrojevi.get(i);
            BigDecimal kol = kolicine.get(i);
            stavke.add(excelReaderService.napraviStavku(rb, kol));
        }
        
        Map<String, String> repl = new HashMap<>();
        repl.put("{{DATUM}}", metadata.getDatumFormatiranDot());
        repl.put("{{DATUM1}}", metadata.getDatumFormatiranSlash());
        repl.put("{{K}}", metadata.getKBroj());
        repl.put("{{BROJ}}", metadata.getBroj());
        repl.put("{{RASKRSNICA}}", metadata.getNazivRaskrsnice());
        
        WordTabelaPopunjavanje.fillMetadata(doc, repl);
        WordTabelaPopunjavanje.fillTable(doc, stavke);
    }
}

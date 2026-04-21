package app;

import domain.ZapisnikMetadata;
import excel.ExcelCenovnik;
import excel.CenovnikLookup;
import repository.RaskrsniceRepository;
import service.ExcelReaderService;
import service.ZapisnikService;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import word.WordTabelaPopunjavanje;

import javax.print.PrintServiceLookup;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.math.BigDecimal;
import java.nio.file.Path;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.List;
import java.util.Scanner;
import java.awt.Desktop;
import java.io.File;

public class ZapisniciApp {
    
    private static final Path RASKRSNICE_XLSX = Path.of("src/main/resources/Nazivi_raskrsnica.xlsx");
    private static final Path CENOVNIK_XLSX   = Path.of("src/main/resources/Blanko za izradu zapisnika.xlsx");
    private static final Path TEMPLATE_DOCX   = Path.of("src/main/resources/template.docx");
    private static final Path OUTPUT_DIR      = Path.of("output");
    
    public static void main(String[] args) throws Exception {
        
        try (Scanner sc = new Scanner(System.in)) {
            
            RaskrsniceRepository raskrsniceRepository = new RaskrsniceRepository(RASKRSNICE_XLSX);
            CenovnikLookup cenovnikLookup = new ExcelCenovnik(CENOVNIK_XLSX);
            
            ExcelReaderService excelReaderService = new ExcelReaderService(cenovnikLookup, raskrsniceRepository);
            ZapisnikService zapisnikService = new ZapisnikService(excelReaderService);
            BigDecimal ukupanZbirPrograma = BigDecimal.ZERO;
            
            while (true) {
                LocalDate datum = readDate(sc, "Unesi datum (dd.MM.yyyy.): ", "dd.MM.yyyy.");
                System.out.print("Unesi K broj: ");
                String kBroj = "K" + sc.nextLine().trim();
                
                System.out.print("Unesi broj zapisnika: ");
                String broj = sc.nextLine().trim();
                
                String nazivRaskrsnice = raskrsniceRepository.getNazivRaskrsnice(kBroj);
                if (nazivRaskrsnice == null || nazivRaskrsnice.isBlank()) {
                    System.out.println("Ne postoji raskrsnica za K broj: " + kBroj);
                    if (!askForAnother(sc)) break;
                    continue;
                }
                
                ZapisnikMetadata metadata = new ZapisnikMetadata(datum, kBroj, broj, nazivRaskrsnice);
                
                List<Integer> redniBrojevi = new ArrayList<>();
                List<BigDecimal> kolicine = new ArrayList<>();
                
                while (true) {
                    int rb = readInt(sc, "Unesi R.br (0 za kraj): ");
                    if (rb == 0) break;
                    
                    BigDecimal kolicina = readBigDecimal(sc, "Unesi kolicinu: ");
                    redniBrojevi.add(rb);
                    kolicine.add(kolicina);
                }
                
                if (redniBrojevi.isEmpty()) {
                    System.out.println("Nisi uneo nijednu stavku. Preskacem generisanje.");
                    if (!askForAnother(sc)) break;
                    continue;
                }
                
                String outName = buildOutputFileName(metadata);
                Path outPath = OUTPUT_DIR.resolve(outName);
                
                try (FileInputStream in = new FileInputStream(TEMPLATE_DOCX.toFile());
                     XWPFDocument doc = new XWPFDocument(in)) {
                    
                    BigDecimal zbirUkupno = (zapisnikService.generisiZapisnik(doc, metadata, redniBrojevi, kolicine));
                    ukupanZbirPrograma =ukupanZbirPrograma.add(zbirUkupno);
                    
                    OUTPUT_DIR.toFile().mkdirs();
                    try (FileOutputStream out = new FileOutputStream(outPath.toFile())) {
                        doc.write(out);
                    }
                }
                
                System.out.println("Zapisnik sacuvan: " + outPath);
                
                if (askForPrint(sc)) {
                    try {
                        printWordDocument(outPath);
                        System.out.println("Dokument je poslat na podrazumevani stampac.");
                    } catch (Exception e) {
                        System.out.println("Stampanje nije uspelo: " + e.getMessage());
                    }
                }
                
                if (!askForAnother(sc)) {
                    System.out.println("Ukupan zbir: "+ WordTabelaPopunjavanje.formatMoney(ukupanZbirPrograma));
                    break;
                }
            }
        }
    }
    
    private static boolean askForAnother(Scanner sc) {
        System.out.print("Da li zelis da napravis jos jedan zapisnik? 1-Da 2-Ne: ");
        String s = sc.nextLine().trim();
        return "1".equals(s);
    }
    
    private static boolean askForPrint(Scanner sc) {
        System.out.print("Da li zelis da odstampas zapisnik na podrazumevani stampac? 1-Da 2-Ne: ");
        String s = sc.nextLine().trim();
        return "1".equals(s);
    }
    
    private static LocalDate readDate(Scanner sc, String prompt, String pattern) {
        DateTimeFormatter fmt = DateTimeFormatter.ofPattern(pattern);
        while (true) {
            System.out.print(prompt);
            String input = sc.nextLine().trim();
            try {
                return LocalDate.parse(input, fmt);
            } catch (Exception e) {
                System.out.println("Neispravan datum. Ocekujem format: " + pattern);
            }
        }
    }
    
    private static int readInt(Scanner sc, String prompt) {
        while (true) {
            System.out.print(prompt);
            String s = sc.nextLine().trim();
            try {
                return Integer.parseInt(s);
            } catch (Exception e) {
                System.out.println("Neispravan broj. Pokusaj ponovo.");
            }
        }
    }
    
    private static BigDecimal readBigDecimal(Scanner sc, String prompt) {
        while (true) {
            System.out.print(prompt);
            String s = sc.nextLine().trim().replace(",", ".");
            try {
                return new BigDecimal(s);
            } catch (Exception e) {
                System.out.println("Neispravna decimalna vrednost. Primer: 1.5 ili 1,5");
            }
        }
    }
    
    private static String buildOutputFileName(ZapisnikMetadata m) {
        return m.getBroj()+".docx";
    }
    
    private static void printWordDocument(Path documentPath) throws Exception {
        if (documentPath == null) {
            throw new IllegalArgumentException("Putanja do dokumenta ne sme biti null.");
        }
        
        if (!Desktop.isDesktopSupported()) {
            throw new UnsupportedOperationException("Desktop API nije podrzan na ovom sistemu.");
        }
        
        Desktop desktop = Desktop.getDesktop();
        
        if (!desktop.isSupported(Desktop.Action.PRINT)) {
            throw new UnsupportedOperationException("PRINT akcija nije podrzana na ovom sistemu.");
        }
        
        File file = documentPath.toFile();
        if (!file.exists()) {
            throw new IllegalStateException("Dokument za stampu ne postoji: " + documentPath);
        }
        
        desktop.print(file);
    }
}

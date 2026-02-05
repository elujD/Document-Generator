import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTText;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTR;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.math.BigDecimal;
import java.nio.file.Path;
import java.sql.SQLOutput;
import java.util.*;
import java.util.concurrent.ConcurrentHashMap;

public class ExcelReaderService {
    
    public static void main(String[] args) throws Exception {
        Scanner scanner = new Scanner(System.in);
        int odluka=0;
        
        // ====== Raskrsnice ======
        Path raskrsnicePath = Path.of("src/main/resources/Nazivi_raskrsnica.xlsx");
        RaskrsniceRepository repo = new RaskrsniceRepository(raskrsnicePath);
        
        // ====== Cenovnik ======
        Path cenovnikPath = Path.of("src/main/resources/Blanko za izradu zapisnika.xlsx");
        CenovnikRepository cenovnik = new CenovnikRepository(cenovnikPath);
        
        while(true){
            List<Stavka> stavke = new ArrayList<>();
            
            System.out.print("Unesi datum: ");
            String datum = scanner.nextLine();
            
            System.out.print("Unesi K: ");
            String kBroj = "K" + scanner.nextLine();
            
            System.out.print("Unesi Broj: ");
            String broj = scanner.nextLine() + "-" + datum.substring(datum.length() - 2);
            
            String naziv = repo.getNazivRaskrsnice(kBroj);
            String datum1 = datum.replace(".", "/");
            
            while (true) {
                System.out.print("Unesi RB stavke (0 za kraj): ");
                int rb = parseIntSafe(scanner.nextLine());
                if (rb == 0) break;
                
                System.out.print("Unesi kolicinu: ");
                BigDecimal kol = parseBigDecimalSafe(scanner.nextLine());
                if (kol == null) {
                    System.out.println("Neispravna kolicina.");
                    continue;
                }
                
                CenovnikRow row = cenovnik.getByRb(rb);
                if (row == null) {
                    System.out.println("Ne postoji stavka sa RB = " + rb + " u cenovniku.");
                    continue;
                }
                
                stavke.add(new Stavka(rb, row.modulUsluga, row.jedinicaMere, kol, row.cena));
            }
            
            // ====== Word ======
            String templatePath = "src/main/resources/template.docx";
            
            try (FileInputStream fis = new FileInputStream(templatePath);
                 XWPFDocument doc = new XWPFDocument(fis)) {
                
                // 1) Replace global placeholder-a (cuva format)
                Map<String, String> repl = new HashMap<>();
                repl.put("{{DATUM}}", datum);
                repl.put("{{K}}", kBroj);
                repl.put("{{DATUM1}}", datum1);
                repl.put("{{RASKRSNICA}}", naziv);
                repl.put("{{BROJ}}", broj);
                
                replaceAllPlaceholdersPreserveFormatting(doc, repl);
                
                // 2) Popuni tabelu stavkama + dodaj SUM red (cuva format template reda)
                fillStavkeTablePreserveFormatting(doc, stavke);
                
                // 3) Snimi
                try (FileOutputStream out = new FileOutputStream(broj + ".docx")) {
                    doc.write(out);
                }
            }
            
            System.out.println("Dokument je uspesno kreiran.");
            System.out.println("Da li zelis dalje da pravis dalje ? 1-Da, 2-Ne");
            odluka=parseIntSafe(scanner.nextLine());
            if(odluka==2)break;
        }
        System.out.println("KRAJ ZAPISNIKA!");
    }
    
    // ============================================================
    //  PLACEHOLDER REPLACE (cuva format)
    // ============================================================
    private static void replaceAllPlaceholdersPreserveFormatting(XWPFDocument doc, Map<String, String> repl) {
        for (XWPFParagraph p : doc.getParagraphs()) {
            replaceInParagraphPreserveRuns(p, repl);
        }
        for (XWPFTable t : doc.getTables()) {
            for (XWPFTableRow r : t.getRows()) {
                for (XWPFTableCell c : r.getTableCells()) {
                    for (XWPFParagraph p : c.getParagraphs()) {
                        replaceInParagraphPreserveRuns(p, repl);
                    }
                }
            }
        }
    }
    
    private static void replaceInParagraphPreserveRuns(XWPFParagraph p, Map<String, String> repl) {
        if (p == null) return;
        List<XWPFRun> runs = p.getRuns();
        if (runs == null || runs.isEmpty()) return;
        
        // 1) pokušaj run-by-run
        for (XWPFRun run : runs) {
            String text = getRunFullText(run);
            if (text == null || text.isEmpty()) continue;
            
            String newText = text;
            for (Map.Entry<String, String> e : repl.entrySet()) {
                String key = e.getKey();
                String val = e.getValue();
                if (val == null) val = ""; // ili "N/A"
                newText = newText.replace(key, val);
            }
            if (!newText.equals(text)) {
                setRunFullText(run, newText);
            }
        }
        
        // 2) fallback: ceo paragraf (split placeholder preko vise run-ova)
        String full = p.getText();
        if (full == null) return;
        
        String replaced = full;
        for (Map.Entry<String, String> e : repl.entrySet()) {
            String key = e.getKey();
            String val = e.getValue();
            if (val == null) val = "";
            replaced = replaced.replace(key, val);
        }
        
        if (!replaced.equals(full)) {
            setRunFullText(runs.get(0), replaced);
            for (int i = 1; i < runs.size(); i++) {
                setRunFullText(runs.get(i), "");
            }
        }
    }
    
    // ============================================================
    //  TABLE FILL (cuva format template reda) + SUM
    // ============================================================
    private static void fillStavkeTablePreserveFormatting(XWPFDocument doc, List<Stavka> stavke) {
        XWPFTable table = null;
        int templateRowIndex = -1;
        
        // Nadji template red koji sadrzi {{STAVKE}}
        for (XWPFTable t : doc.getTables()) {
            for (int i = 0; i < t.getNumberOfRows(); i++) {
                if (rowContains(t.getRow(i), "{{STAVKE}}")) {
                    table = t;
                    templateRowIndex = i;
                    break;
                }
            }
            if (table != null) break;
        }
        
        if (table == null) {
            throw new IllegalStateException("Template red sa {{STAVKE}} nije pronađen u Word tabeli.");
        }
        
        XWPFTableRow templateRow = table.getRow(templateRowIndex);
        int insertIndex = templateRowIndex + 1;
        
        BigDecimal sumaUkupno = BigDecimal.ZERO;
        
        // 1) Ubaci stavke
        for (Stavka s : stavke) {
            XWPFTableRow newRow = table.insertNewTableRow(insertIndex++);
            cloneRowKeepingFormatting_SAFE(templateRow, newRow);
            
            Map<String, String> rep = new HashMap<>();
            rep.put("{{STAVKE}}", "");
            rep.put("{{RB}}", String.valueOf(s.rb));
            rep.put("{{MODUL}}", s.modulUsluga);
            rep.put("{{KOL}}", toPlain(s.kolicina));
            rep.put("{{CENA}}", toPlain(s.cena));
            rep.put("{{UKUPNO}}", toPlain(s.ukupno));
            
            replaceInRowParagraphs(newRow, rep);
            
            sumaUkupno = sumaUkupno.add(s.ukupno);
        }
        
        // 2) SUM red na dnu
        XWPFTableRow sumRow = table.createRow();
        cloneRowKeepingFormatting_SAFE(templateRow, sumRow);
        
        // očisti placeholder-e u SUM redu
        Map<String, String> clear = new HashMap<>();
        clear.put("{{STAVKE}}", "");
        clear.put("{{RB}}", "");
        clear.put("{{MODUL}}", "");
        clear.put("{{KOL}}", "");
        clear.put("{{CENA}}", "");
        clear.put("{{UKUPNO}}", "");
        replaceInRowParagraphs(sumRow, clear);
        
        // upiši "UKUPNO:" u pretposlednju kolonu (CENA = index 4) i zbir u poslednju (UKUPNO = index 5)
        // Ako ti je raspored kolona drugačiji, promeni indekse.
        setCellParagraphTextPreserveFirstRun(sumRow.getCell(4), "UKUPNO:");
        setCellParagraphTextPreserveFirstRun(sumRow.getCell(5), toPlain(sumaUkupno));
        
        // 3) obriši template red
        table.removeRow(templateRowIndex);
    }
    
    private static boolean rowContains(XWPFTableRow row, String token) {
        if (row == null) return false;
        for (XWPFTableCell cell : row.getTableCells()) {
            String txt = cell.getText();
            if (txt != null && txt.contains(token)) return true;
        }
        return false;
    }
    
    /**
     * Replace na nivou paragrafa u ćeliji (100% radi i kad je placeholder splitovan na više run-ova),
     * i zadržava stil prvog run-a.
     */
    private static void replaceInRowParagraphs(XWPFTableRow row, Map<String, String> rep) {
        for (XWPFTableCell cell : row.getTableCells()) {
            for (XWPFParagraph p : cell.getParagraphs()) {
                List<XWPFRun> runs = p.getRuns();
                if (runs == null || runs.isEmpty()) continue;
                
                StringBuilder sb = new StringBuilder();
                for (XWPFRun r : runs) sb.append(getRunFullText(r));
                String full = sb.toString();
                if (full.isEmpty()) continue;
                
                String replaced = full;
                for (Map.Entry<String, String> e : rep.entrySet()) {
                    replaced = replaced.replace(e.getKey(), e.getValue());
                }
                
                if (!replaced.equals(full)) {
                    setRunFullText(runs.get(0), replaced);
                    for (int i = 1; i < runs.size(); i++) setRunFullText(runs.get(i), "");
                }
            }
        }
    }
    
    /**
     * Stabilno kloniranje reda preko POI API-ja (bez direktnog XML kopiranja),
     * uz kopiranje TcPr (ćelija) + PPr (paragraf) + RPr (run).
     */
    private static void cloneRowKeepingFormatting_SAFE(XWPFTableRow src, XWPFTableRow dst) {
        if (src.getCtRow().isSetTrPr()) {
            dst.getCtRow().setTrPr(src.getCtRow().getTrPr());
        }
        
        // očisti postojeće ćelije u dst
        int existing = dst.getTableCells().size();
        for (int i = existing - 1; i >= 0; i--) {
            dst.removeCell(i);
        }
        
        int cellCount = src.getTableCells().size();
        
        for (int i = 0; i < cellCount; i++) {
            XWPFTableCell srcCell = src.getCell(i);
            XWPFTableCell dstCell = dst.createCell();
            
            if (srcCell.getCTTc().isSetTcPr()) {
                dstCell.getCTTc().setTcPr(srcCell.getCTTc().getTcPr());
            }
            
            while (dstCell.getParagraphs().size() > 0) dstCell.removeParagraph(0);
            
            for (XWPFParagraph sp : srcCell.getParagraphs()) {
                XWPFParagraph dp = dstCell.addParagraph();
                
                if (sp.getCTP().isSetPPr()) {
                    dp.getCTP().setPPr(sp.getCTP().getPPr());
                }
                
                for (XWPFRun sr : sp.getRuns()) {
                    XWPFRun dr = dp.createRun();
                    
                    CTR srCtr = sr.getCTR();
                    if (srCtr.isSetRPr()) {
                        dr.getCTR().setRPr(srCtr.getRPr());
                    }
                    
                    String t = getRunFullText(sr);
                    if (t != null && !t.isEmpty()) {
                        setRunFullText(dr, t);
                    }
                }
                
                if (dp.getRuns().isEmpty()) dp.createRun();
            }
            
            if (dstCell.getParagraphs().isEmpty()) dstCell.addParagraph().createRun();
        }
    }
    
    private static void setCellParagraphTextPreserveFirstRun(XWPFTableCell cell, String text) {
        if (cell == null) return;
        
        XWPFParagraph p;
        if (cell.getParagraphs() == null || cell.getParagraphs().isEmpty()) {
            p = cell.addParagraph();
        } else {
            p = cell.getParagraphs().get(0);
        }
        
        List<XWPFRun> runs = p.getRuns();
        if (runs == null || runs.isEmpty()) {
            XWPFRun r = p.createRun();
            setRunFullText(r, text);
            return;
        }
        
        setRunFullText(runs.get(0), text);
        for (int i = 1; i < runs.size(); i++) setRunFullText(runs.get(i), "");
    }
    
    // ============================================================
    //  RUN TEXT HELPERS (ne puca na CTR edge-case)
    // ============================================================
    private static String getRunFullText(XWPFRun run) {
        if (run == null) return "";
        try {
            CTR ctr = run.getCTR();
            int n = ctr.sizeOfTArray();
            if (n <= 0) return "";
            StringBuilder sb = new StringBuilder();
            for (int i = 0; i < n; i++) {
                CTText t = ctr.getTArray(i);
                if (t != null && t.getStringValue() != null) sb.append(t.getStringValue());
            }
            return sb.toString();
        } catch (Exception e) {
            return "";
        }
    }
    
    private static void setRunFullText(XWPFRun run, String newText) {
        if (run == null) return;
        try {
            CTR ctr = run.getCTR();
            int n = ctr.sizeOfTArray();
            for (int i = n - 1; i >= 0; i--) ctr.removeT(i);
            ctr.addNewT().setStringValue(newText == null ? "" : newText);
        } catch (Exception ignored) {
            try {
                run.setText(newText == null ? "" : newText, 0);
            } catch (Exception ignored2) {
            }
        }
    }
    
    // ============================================================
    //  PARSE HELPERS
    // ============================================================
    private static int parseIntSafe(String s) {
        try {
            return Integer.parseInt(s.trim());
        } catch (Exception e) {
            return 0;
        }
    }
    
    private static BigDecimal parseBigDecimalSafe(String s) {
        if (s == null) return null;
        try {
            String v = s.trim().replace(".", "").replace(",", ".");
            if (v.isEmpty()) return null;
            return new BigDecimal(v);
        } catch (Exception e) {
            return null;
        }
    }
    
    private static String toPlain(BigDecimal v) {
        if (v == null) return "";
        return v.stripTrailingZeros().toPlainString();
    }
    
    // ============================================================
    //  DATA CLASSES + EXCEL CENOVNIK (BigDecimal cena)
    // ============================================================
    public static class Stavka {
        public final int rb;
        public final String modulUsluga;
        public final String jedinicaMere;
        public final BigDecimal kolicina;
        public final BigDecimal cena;
        public final BigDecimal ukupno;
        
        public Stavka(int rb, String modulUsluga, String jedinicaMere, BigDecimal kolicina, BigDecimal cena) {
            this.rb = rb;
            this.modulUsluga = modulUsluga;
            this.jedinicaMere = jedinicaMere;
            this.kolicina = kolicina;
            this.cena = cena;
            this.ukupno = kolicina.multiply(cena);
        }
    }
    
    public static class CenovnikRow {
        public final int rb;
        public final String modulUsluga;
        public final String jedinicaMere;
        public final BigDecimal cena;
        
        public CenovnikRow(int rb, String modulUsluga, String jedinicaMere, BigDecimal cena) {
            this.rb = rb;
            this.modulUsluga = modulUsluga;
            this.jedinicaMere = jedinicaMere;
            this.cena = cena;
        }
    }
    
    public static class CenovnikRepository {
        private final Map<Integer, CenovnikRow> byRb = new ConcurrentHashMap<>();
        
        public CenovnikRepository(Path excelPath) {
            load(excelPath);
        }
        
        public CenovnikRow getByRb(int rb) {
            return byRb.get(rb);
        }
        
        private void load(Path excelPath) {
            try (FileInputStream fis = new FileInputStream(excelPath.toFile());
                 Workbook wb = new XSSFWorkbook(fis)) {
                
                Sheet sheet = wb.getSheetAt(0);
                
                // Ako se raspored pomeri, promeni startRow
                int startRow = 3; // 0-based index 3 ~ Excel red 4
                for (int r = startRow; r <= sheet.getLastRowNum(); r++) {
                    Row row = sheet.getRow(r);
                    if (row == null) continue;
                    
                    Integer rb = readInt(row.getCell(0));
                    if (rb == null) continue;
                    
                    String modul = readString(row.getCell(1));
                    if (modul == null || modul.isBlank()) continue;
                    
                    String jm = readString(row.getCell(2));
                    if (jm == null || jm.isBlank()) jm = "kom.";
                    
                    BigDecimal cena = readBigDecimal(row.getCell(4));
                    if (cena == null) continue;
                    
                    byRb.put(rb, new CenovnikRow(rb, modul.trim(), jm.trim(), cena));
                }
            } catch (Exception e) {
                throw new RuntimeException("Greska pri citanju cenovnika iz Excela: " + e.getMessage(), e);
            }
        }
        
        private static String readString(Cell cell) {
            if (cell == null) return null;
            try {
                if (cell.getCellType() == CellType.STRING) return cell.getStringCellValue();
                if (cell.getCellType() == CellType.NUMERIC) return String.valueOf(cell.getNumericCellValue());
                if (cell.getCellType() == CellType.FORMULA) {
                    try {
                        return cell.getStringCellValue();
                    } catch (Exception ignored) {
                        return String.valueOf(cell.getNumericCellValue());
                    }
                }
            } catch (Exception ignored) {
            }
            return null;
        }
        
        private static Integer readInt(Cell cell) {
            if (cell == null) return null;
            try {
                if (cell.getCellType() == CellType.NUMERIC) return (int) Math.round(cell.getNumericCellValue());
                if (cell.getCellType() == CellType.STRING) {
                    String s = cell.getStringCellValue();
                    if (s == null) return null;
                    s = s.trim();
                    if (s.isEmpty()) return null;
                    s = s.replace(".", "").replace(",", "");
                    return Integer.parseInt(s);
                }
                if (cell.getCellType() == CellType.FORMULA) return (int) Math.round(cell.getNumericCellValue());
            } catch (Exception ignored) {
                return null;
            }
            return null;
        }
        
        private static BigDecimal readBigDecimal(Cell cell) {
            if (cell == null) return null;
            try {
                if (cell.getCellType() == CellType.NUMERIC) {
                    return BigDecimal.valueOf(cell.getNumericCellValue());
                }
                if (cell.getCellType() == CellType.STRING) {
                    String s = cell.getStringCellValue();
                    if (s == null) return null;
                    s = s.trim();
                    if (s.isEmpty()) return null;
                    
                    // "1.234,56" -> "1234.56"
                    s = s.replace(".", "").replace(",", ".");
                    return new BigDecimal(s);
                }
                if (cell.getCellType() == CellType.FORMULA) {
                    return BigDecimal.valueOf(cell.getNumericCellValue());
                }
            } catch (Exception ignored) {
                return null;
            }
            return null;
        }
    }
}

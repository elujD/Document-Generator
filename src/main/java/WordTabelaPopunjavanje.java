import org.apache.poi.xwpf.usermodel.*;

import java.util.List;

public class WordTabelaPopunjavanje {
    
    public static void fillTable(XWPFDocument doc, List<Stavka> stavke) {
        XWPFTable table = null;
        int templateRowIndex = -1;
        
        // 1) Nađi tabelu + red koji sadrži {{STAVKE}}
        for (XWPFTable t : doc.getTables()) {
            for (int i = 0; i < t.getNumberOfRows(); i++) {
                XWPFTableRow row = t.getRow(i);
                if (rowContains(row, "{{STAVKE}}")) {
                    table = t;
                    templateRowIndex = i;
                    break;
                }
            }
            if (table != null) break;
        }
        
        if (table == null) {
            throw new RuntimeException("Template red sa {{STAVKE}} nije pronađen u Word tabeli.");
        }
        
        XWPFTableRow templateRow = table.getRow(templateRowIndex);
        int insertIndex = templateRowIndex + 1;
        
        // 2) Ubaci redove
        for (Stavka s : stavke) {
            XWPFTableRow newRow = table.insertNewTableRow(insertIndex++);
            copyRow(templateRow, newRow);
            
            // 0: R.br, 1: Modul/Usluga, 2: Jed. mere, 3: Količina, 4: Cena, 5: Ukupno
            set(newRow.getCell(0), String.valueOf(s.rb));
            set(newRow.getCell(1), s.modulUsluga);
            set(newRow.getCell(2), s.jedinicaMere);
            set(newRow.getCell(3), String.valueOf(s.kolicina));
            set(newRow.getCell(4), String.valueOf(s.cena));
            set(newRow.getCell(5), String.valueOf(s.ukupno));
        }
        
        // 3) Obriši template red
        table.removeRow(templateRowIndex);
    }
    
    private static boolean rowContains(XWPFTableRow row, String token) {
        if (row == null) return false;
        for (XWPFTableCell cell : row.getTableCells()) {
            String text = cell.getText(); // ovo postoji
            if (text != null && text.contains(token)) return true;
        }
        return false;
    }
    
    private static void copyRow(XWPFTableRow src, XWPFTableRow dst) {
        // kopiraj row properties (opciono)
        if (src.getCtRow().isSetTrPr()) {
            dst.getCtRow().setTrPr(src.getCtRow().getTrPr());
        }
        
        // dst red po default-u često nema nijednu ćeliju → koristimo createCell()
        int cells = src.getTableCells().size();
        for (int i = 0; i < cells; i++) {
            XWPFTableCell srcCell = src.getCell(i);
            XWPFTableCell dstCell = dst.createCell();
            
            // kopiraj cell properties (širine, margine, borderi...)
            if (srcCell.getCTTc().isSetTcPr()) {
                dstCell.getCTTc().setTcPr(srcCell.getCTTc().getTcPr());
            }
            
            // očisti postojeće paragraf-e i ostavi jedan
            while (dstCell.getParagraphs().size() > 0) dstCell.removeParagraph(0);
            dstCell.addParagraph().createRun();
        }
        
        // POI zna da napravi jedan višak cell na početku u nekim verzijama,
        // ali ovde kreiramo tačno koliko treba, pa je OK.
    }
    
    private static void set(XWPFTableCell cell, String value) {
        // obriši sve paragraf-e i upiši jedan
        while (cell.getParagraphs().size() > 0) cell.removeParagraph(0);
        XWPFParagraph p = cell.addParagraph();
        p.createRun().setText(value == null ? "" : value);
    }
}

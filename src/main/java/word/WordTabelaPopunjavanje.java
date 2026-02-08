package word;

import domain.Stavka;
import org.apache.poi.xwpf.usermodel.*;

import java.math.BigDecimal;
import java.util.List;
import java.text.DecimalFormat;
import java.text.DecimalFormatSymbols;
import java.util.Locale;
import java.math.RoundingMode;

public class WordTabelaPopunjavanje {
    
    private static final String STAVKE_TOKEN = "{{STAVKE}}";
    
    public static void fillTable(XWPFDocument doc, List<Stavka> stavke) {
        if (doc == null) throw new IllegalArgumentException("doc ne sme biti null.");
        if (stavke == null) throw new IllegalArgumentException("stavke ne smeju biti null.");
        
        XWPFTable table = null;
        int templateRowIndex = -1;
        
        outer:
        for (XWPFTable t : doc.getTables()) {
            for (int i = 0; i < t.getNumberOfRows(); i++) {
                XWPFTableRow row = t.getRow(i);
                if (rowContains(row, STAVKE_TOKEN)) {
                    table = t;
                    templateRowIndex = i;
                    break outer;
                }
            }
        }
        
        if (table == null) {
            throw new IllegalStateException("Template red sa {{STAVKE}} nije pronađen u Word tabeli.");
        }
        
        XWPFTableRow templateRow = table.getRow(templateRowIndex);
        int insertIndex = templateRowIndex + 1;
        
        for (Stavka s : stavke) {
            XWPFTableRow newRow = table.insertNewTableRow(insertIndex++);
            copyRowStructureAndStyle(templateRow, newRow);
            
            setTextPreserveStyle(newRow.getCell(0), String.valueOf(s.rb));
            setTextPreserveStyle(newRow.getCell(1), safe(s.modulUsluga));
            setTextPreserveStyle(newRow.getCell(2), safe(s.jedinicaMere));
            setTextPreserveStyle(newRow.getCell(3), bd(s.kolicina));
            setTextPreserveStyle(newRow.getCell(4), formatMoney(s.cena));
            setTextPreserveStyle(newRow.getCell(5), formatMoney(s.ukupno));
        }
        
        addUkupnoRow(table,stavke);
        table.removeRow(templateRowIndex);
    }
    
    private static boolean rowContains(XWPFTableRow row, String token) {
        if (row == null) return false;
        for (XWPFTableCell cell : row.getTableCells()) {
            String text = cell.getText();
            if (text != null && text.contains(token)) return true;
        }
        return false;
    }
    
    private static void copyRowStructureAndStyle(XWPFTableRow src, XWPFTableRow dst) {
        
        if (src.getCtRow().isSetTrPr()) {
            dst.getCtRow().setTrPr(src.getCtRow().getTrPr());
        }
        
        int cells = src.getTableCells().size();
        
        for (int i = 0; i < cells; i++) {
            XWPFTableCell srcCell = src.getCell(i);
            XWPFTableCell dstCell = (i == 0 && dst.getTableCells().size() > 0)
                    ? dst.getCell(0)
                    : dst.createCell();
            
            if (srcCell.getCTTc().isSetTcPr()) {
                dstCell.getCTTc().setTcPr(srcCell.getCTTc().getTcPr());
            }
            
            while (dstCell.getParagraphs().size() > 0) dstCell.removeParagraph(0);
            
            XWPFParagraph srcP = srcCell.getParagraphs().isEmpty() ? null : srcCell.getParagraphs().get(0);
            XWPFParagraph dstP = dstCell.addParagraph();
            
            if (srcP != null && srcP.getCTP().isSetPPr()) {
                dstP.getCTP().setPPr(srcP.getCTP().getPPr());
            }
            
            XWPFRun dstRun = dstP.createRun();
            XWPFRun srcRun = (srcP != null && !srcP.getRuns().isEmpty()) ? srcP.getRuns().get(0) : null;
            
            if (srcRun != null && srcRun.getCTR().isSetRPr()) {
                dstRun.getCTR().setRPr(srcRun.getCTR().getRPr());
            }
        }
    }
    
    private static void setTextPreserveStyle(XWPFTableCell cell, String value) {
        if (cell == null) return;
        
        XWPFParagraph p;
        if (cell.getParagraphs().isEmpty()) {
            p = cell.addParagraph();
        } else {
            p = cell.getParagraphs().get(0);
        }
        
        XWPFRun run;
        if (p.getRuns().isEmpty()) {
            run = p.createRun();
        } else {
            run = p.getRuns().get(0);
        }
        
        run.setText("", 0);
        run.setText(value == null ? "" : value, 0);
    }
    
    private static String safe(String s) {
        return s == null ? "" : s;
    }
    
    private static String bd(BigDecimal x) {
        return x == null ? "" : x.stripTrailingZeros().toPlainString();
    }
    
    public static void fillMetadata(XWPFDocument doc, java.util.Map<String, String> replacements) {
        if (doc == null) throw new IllegalArgumentException("doc ne sme biti null.");
        if (replacements == null) throw new IllegalArgumentException("replacements ne sme biti null.");
        
        for (XWPFParagraph p : doc.getParagraphs()) {
            replaceInParagraph(p, replacements);
        }
        
        for (XWPFTable t : doc.getTables()) {
            for (XWPFTableRow r : t.getRows()) {
                for (XWPFTableCell c : r.getTableCells()) {
                    for (XWPFParagraph p : c.getParagraphs()) {
                        replaceInParagraph(p, replacements);
                    }
                }
            }
        }
    }
    
    private static void replaceInParagraph(XWPFParagraph paragraph, java.util.Map<String, String> replacements) {
        if (paragraph == null) return;
        
        String fullText = paragraph.getText();
        if (fullText == null || fullText.isEmpty()) return;
        
        String replaced = fullText;
        for (var e : replacements.entrySet()) {
            String key = e.getKey();
            String val = e.getValue() == null ? "" : e.getValue();
            replaced = replaced.replace(key, val);
        }
        
        if (replaced.equals(fullText)) return;
        
        java.util.List<XWPFRun> runs = paragraph.getRuns();
        if (runs == null || runs.isEmpty()) {
            XWPFRun r = paragraph.createRun();
            r.setText(replaced, 0);
            return;
        }
        
        XWPFRun first = runs.get(0);
        
        for (int i = runs.size() - 1; i >= 1; i--) {
            paragraph.removeRun(i);
        }
        
        first.setText("", 0);
        first.setText(replaced, 0);
    }
    
    private static void addUkupnoRow(XWPFTable table, List<Stavka> stavke) {
        BigDecimal ukupno = BigDecimal.ZERO;
        
        for (Stavka s : stavke) {
            ukupno = ukupno.add(s.getUkupno());
        }
        
        XWPFTableRow totalRow = table.createRow();
        
        XWPFTableCell ukupnoLabelCell = totalRow.getCell(4);
        ukupnoLabelCell.removeParagraph(0);
        
        XWPFParagraph labelParagraph = ukupnoLabelCell.addParagraph();
        labelParagraph.setAlignment(ParagraphAlignment.LEFT);
        labelParagraph.createRun().setText("UKUPNO");
        
        totalRow.getCell(0).setText("");
        totalRow.getCell(1).setText("");
        totalRow.getCell(2).setText("");
        totalRow.getCell(3).setText("");
        
        XWPFTableCell totalCell = totalRow.getCell(5);
        totalCell.removeParagraph(0);
        
        XWPFParagraph totalParagraph = totalCell.addParagraph();
        totalParagraph.setAlignment(ParagraphAlignment.RIGHT);
        totalParagraph.createRun().setText(formatMoney(ukupno));
    }
    
    private static final DecimalFormat MONEY_FORMAT;
    static {
        DecimalFormatSymbols symbols = new DecimalFormatSymbols(Locale.US);
        MONEY_FORMAT = new DecimalFormat("#,##0.00", symbols);
        MONEY_FORMAT.setRoundingMode(RoundingMode.HALF_UP);
    }
    
    private static String formatMoney(BigDecimal v) {
        if (v == null) return "";
        return MONEY_FORMAT.format(v);
    }
}

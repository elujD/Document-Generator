package excel;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.math.BigDecimal;
import java.util.HashMap;
import java.util.Map;

public class ExcelCenovnik {
    
    public static class CenovnikRow {
        public final int rb;
        public final String modulUsluga;
        public final String jedinicaMere;
        public final BigDecimal cena;
        
        public CenovnikRow(int rb,
                           String modulUsluga,
                           String jedinicaMere,
                           BigDecimal cena) {
            this.rb = rb;
            this.modulUsluga = modulUsluga;
            this.jedinicaMere = jedinicaMere;
            this.cena = cena;
        }
    }
    
    static BigDecimal readBigDecimal(Cell cell) {
        if (cell == null) return null;
        
        try {
            if (cell.getCellType() == CellType.NUMERIC) {
                return BigDecimal.valueOf(cell.getNumericCellValue());
            }
            if (cell.getCellType() == CellType.STRING) {
                String s = cell.getStringCellValue()
                        .trim()
                        .replace(".", "")
                        .replace(",", ".");
                if (s.isEmpty()) return null;
                return new BigDecimal(s);
            }
            if (cell.getCellType() == CellType.FORMULA) {
                return BigDecimal.valueOf(cell.getNumericCellValue());
            }
        } catch (Exception ignored) { }
        
        return null;
    }
    
    private final Map<Integer, CenovnikRow> map = new HashMap<>();
    
    public ExcelCenovnik(String path) {
        load(path);
    }
    
    public CenovnikRow get(int rb) {
        return map.get(rb);
    }
    
    private void load(String path) {
        try (Workbook wb = new XSSFWorkbook(new FileInputStream(path))) {
            Sheet sheet = wb.getSheetAt(0);
            
            // podaci od reda 4 (index 3)
            for (int i = 3; i <= sheet.getLastRowNum(); i++) {
                Row r = sheet.getRow(i);
                if (r == null) continue;
                
                Cell cRb = r.getCell(0);
                Cell cModul = r.getCell(1);
                Cell cJm = r.getCell(2);
                Cell cCena = r.getCell(4);
                
                if (cRb == null || cCena == null) continue;
                
                int rb = (int) cRb.getNumericCellValue();
                String modul = cModul.getStringCellValue();
                String jm = cJm != null ? cJm.getStringCellValue() : "kom.";
                BigDecimal cena = BigDecimal.valueOf((int) cCena.getNumericCellValue());
                
                map.put(rb, new CenovnikRow(rb, modul, jm, cena));
            }
        } catch (Exception e) {
            throw new RuntimeException("Greška pri čitanju Excel fajla", e);
        }
    }
}

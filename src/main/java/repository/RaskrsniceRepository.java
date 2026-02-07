package repository;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Collections;
import java.util.HashMap;
import java.util.Map;
import java.util.Objects;

//Ucitava mapu broj raskrsnice -> naziv raskrsnice

public class RaskrsniceRepository {
    
    private static final String PREFERRED_SHEET_NAME = "Final";
    private static final int HEADER_ROW_INDEX = 0;
    
    private static final int CODE_COL_INDEX = 1;
    private static final int NAME_COL_INDEX = 2;
    
    private final Map<String, String> sifraToNaziv;
    
    public RaskrsniceRepository(Path excelPath) throws IOException {
        Objects.requireNonNull(excelPath, "excelPath ne sme biti null");
        this.sifraToNaziv = Collections.unmodifiableMap(loadMap(excelPath));
    }
    
    public String getNazivRaskrsnice(String sifra) {
        String normalized = normalizeCode(sifra);
        if (normalized == null) return null;
        return sifraToNaziv.get(normalized);
    }
    
    private static String normalizeCode(String code) {
        if (code == null) return null;
        String trimmed = code.trim();
        if (trimmed.isEmpty()) return null;
        return trimmed.toUpperCase();
    }
    
    private static Map<String, String> loadMap(Path excelPath) throws IOException {
        Map<String, String> map = new HashMap<>();
        
        try (InputStream is = Files.newInputStream(excelPath);
             Workbook wb = new XSSFWorkbook(is)) {
            
            Sheet sheet = wb.getSheet(PREFERRED_SHEET_NAME);
            if (sheet == null) sheet = wb.getSheetAt(0);
            
            DataFormatter formatter = new DataFormatter();
            
            for (int r = HEADER_ROW_INDEX+1; r <= sheet.getLastRowNum(); r++) {
                Row row = sheet.getRow(r);
                if (row == null) continue;
                
                String rawCode = formatter.formatCellValue(row.getCell(CODE_COL_INDEX));
                String rawName = formatter.formatCellValue(row.getCell(NAME_COL_INDEX));
                
                String code=normalizeCode(rawCode);
                if(code==null) continue;
                
                String name = (rawName == null) ? "" : rawName.trim();
                map.put(code, name);
            }
        }
        return map;
    }
}


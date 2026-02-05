import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.HashMap;
import java.util.Map;

public class RaskrsniceRepository {
    
    private final Map<String, String> sifraToNaziv;
    
    public RaskrsniceRepository(Path excelPath) throws IOException {
        this.sifraToNaziv = ucitajMapu(excelPath);
    }
    
    public String getNazivRaskrsnice(String sifra) {
        if (sifra == null) return null;
        return sifraToNaziv.get(sifra.trim().toUpperCase());
    }
    
    private Map<String, String> ucitajMapu(Path excelPath) throws IOException {
        Map<String, String> map = new HashMap<>();
        
        try (InputStream is = Files.newInputStream(excelPath);
             Workbook wb = new XSSFWorkbook(is)) {
            
            Sheet sheet = wb.getSheet("Final");
            if (sheet == null) sheet = wb.getSheetAt(0);
            
            DataFormatter formatter = new DataFormatter();
            
            // preskacemo header (red 0): Name_iD | IP_Scala | Name | Municipality
            for (int r = 1; r <= sheet.getLastRowNum(); r++) {
                Row row = sheet.getRow(r);
                if (row == null) continue;
                
                String code = formatter.formatCellValue(row.getCell(1)); // kolona B (index 1)
                String name = formatter.formatCellValue(row.getCell(2)); // kolona C (index 2)
                
                if (code != null && !code.isBlank()) {
                    map.put(code.trim().toUpperCase(), name != null ? name.trim() : "");
                }
            }
        }
        
        return map;
    }
}


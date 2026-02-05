import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.List;
import java.util.Scanner;

public class StavkeConsoleInput {
    
    public static List<Stavka> read(ExcelCenovnik cenovnik, Scanner sc) {
        List<Stavka> stavke = new ArrayList<>();
        
        while (true) {
            System.out.print("Unesi R.br (0 za kraj): ");
            String rbStr = sc.nextLine().trim();
            if (rbStr.isEmpty()) continue;
            
            int rb;
            try {
                rb = Integer.parseInt(rbStr);
            } catch (Exception e) {
                System.out.println("Neispravan RB.");
                continue;
            }
            
            if (rb == 0) break;
            
            System.out.print("Unesi količinu: ");
            BigDecimal kolicina;
            try {
                String s = sc.nextLine().trim().replace(",", ".");
                kolicina = new BigDecimal(s);
            } catch (Exception e) {
                System.out.println("Neispravna količina.");
                continue;
            }
            
            ExcelCenovnik.CenovnikRow row = cenovnik.get(rb);
            if (row == null) {
                System.out.println("Ne postoji stavka sa R.br = " + rb);
                continue;
            }
            
            // row.cena mora biti BigDecimal
            stavke.add(new Stavka(
                    rb,
                    row.modulUsluga,
                    row.jedinicaMere,
                    kolicina,
                    row.cena
            ));
        }
        
        return stavke;
    }
}

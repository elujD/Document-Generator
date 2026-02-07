package input;

import domain.Stavka;
import excel.ExcelCenovnik;

import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;
import java.util.Scanner;

public final class StavkeConsoleInput {
    
    public static List<Stavka> read(ExcelCenovnik cenovnik, Scanner sc) {
        Objects.requireNonNull(cenovnik, "Cenovnik ne sme biti null");
        Objects.requireNonNull(sc, "Scanner ne sme biti null");
        
        List<Stavka> stavke = new ArrayList<>();
        
        while (true) {
            Integer rb = readRb(sc);
            
            if (rb == null) continue;
            if (rb == 0) break;
            
            BigDecimal kolicina = readKolicina(sc);
            if (kolicina == null) continue;
            
            ExcelCenovnik.CenovnikRow row = cenovnik.get(rb);
            if (row == null) {
                System.out.println("Ne postoji stavka sa R.br = " + rb);
                continue;
            }
            
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
    
    private static Integer readRb(Scanner sc) {
        System.out.print("Unesi R.br (0 za kraj): ");
        String rbStr = sc.nextLine().trim();
        
        if (rbStr.isEmpty()) {
            return null;
        }
        
        try {
            return Integer.parseInt(rbStr);
        } catch (NumberFormatException e) {
            System.out.println("Neispravan RB. Unesi ceo broj (npr. 1, 2, 3...).");
            return null;
        }
    }
    
    private static BigDecimal readKolicina(Scanner sc) {
        System.out.print("Unesi količinu: ");
        String s = sc.nextLine().trim();
        
        if (s.isEmpty()) {
            System.out.println("Količina ne sme biti prazna.");
            return null;
        }
        
        s = s.replace(",", ".");
        
        try {
            BigDecimal kolicina = new BigDecimal(s);
            
            if (kolicina.compareTo(BigDecimal.ZERO) <= 0) {
                System.out.println("Količina mora biti veća od 0.");
                return null;
            }
            
            return kolicina;
        } catch (NumberFormatException e) {
            System.out.println("Neispravna količina. Primer: 1 ili 1.5");
            return null;
        }
    }
}

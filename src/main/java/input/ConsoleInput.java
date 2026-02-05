package input;

import domain.Stavka;
import domain.ZapisnikMetadata;
import excel.CenovnikLookup;

import java.math.BigDecimal;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.List;
import java.util.Scanner;

//Klasa zaduzena za unos podataka kroz konzolu. Single Responsibility Principle.

public class ConsoleInput {
    
    private final Scanner scanner;
    private final CenovnikLookup cenovnikLookup;
    
    public ConsoleInput(Scanner scanner, CenovnikLookup cenovnikLookup) {
        this.scanner = scanner;
        this.cenovnikLookup = cenovnikLookup;
    }
    
    public ZapisnikMetadata readMetadata(String nazivRaskrsnice) {
        
        System.out.print("Unesi datum (dd.MM.yyyy.): ");
        String datumInput = scanner.nextLine();
        
        LocalDate datum = LocalDate.parse(
                datumInput,
                DateTimeFormatter.ofPattern("dd.MM.yyyy.")
        );
        
        LocalDate datumSlash = LocalDate.parse(datumInput, DateTimeFormatter.ofPattern("dd/MM/yyyy."));
        
        System.out.print("Unesi K broj: ");
        String kBroj = "K" + scanner.nextLine().trim();
        
        System.out.print("Unesi broj zapisnika: ");
        String broj = scanner.nextLine().trim();
        
        return new ZapisnikMetadata(
                datum,
                datumSlash,
                kBroj,
                broj,
                nazivRaskrsnice
        );
    }
    
    public List<Stavka> readStavke() {
        
        List<Stavka> stavke = new ArrayList<>();
        
        while (true) {
            System.out.print("Unesi redni broj stavke (ili ENTER za kraj): ");
            String rbInput = scanner.nextLine().trim();
            
            if (rbInput.isEmpty()) {
                break;
            }
            
            int redniBroj;
            try{
                redniBroj = Integer.parseInt(rbInput);
            }catch(NumberFormatException e){
                System.out.println("Pogrešan unos rednog broja!");
                continue;
            }
            
            CenovnikLookup.CenovnikItem item = cenovnikLookup.findByRedniBroj(redniBroj).orElse(null);
            
            if(item == null){
                System.out.println("Ne postoji stavka u cenovniku!");
                continue;
            }
            
            BigDecimal kolicina;
            while (true) {
                System.out.print("Unesi količinu: ");
                String kolInput = scanner.nextLine().trim();
                
                try {
                    kolicina = new BigDecimal(kolInput);
                    break;
                } catch (NumberFormatException e) {
                    System.out.println("Količina mora biti broj. Pokušaj ponovo.");
                }
            }
            
            Stavka stavka = new Stavka(
                    redniBroj,
                    item.getNaziv(),
                    item.getJm(),
                    kolicina,
                    item.getCena()
            );
            
            stavke.add(stavka);
        }
        
        return stavke;
    }
    
    public boolean askForAnother() {
        System.out.print("Da li želiš da napraviš još jedan zapisnik? 1-Da 2-Ne: ");
        int odgovor = Integer.parseInt(scanner.nextLine());
        return odgovor==1;
    }

}

package input;

import domain.ZapisnikMetadata;
import excel.CenovnikLookup;

import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
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
    
    public boolean askForAnother() {
        System.out.print("Da li želiš da napraviš još jedan zapisnik? 1-Da 2-Ne: ");
        int odgovor = Integer.parseInt(scanner.nextLine());
        return odgovor==1;
    }

}

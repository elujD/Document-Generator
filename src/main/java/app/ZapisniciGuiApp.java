package app;

import domain.ZapisnikMetadata;
import excel.CenovnikLookup;
import excel.ExcelCenovnik;
import repository.RaskrsniceRepository;
import service.ExcelReaderService;
import service.ZapisnikService;
import word.WordTabelaPopunjavanje;

import org.apache.poi.xwpf.usermodel.XWPFDocument;

import javax.swing.*;
import javax.swing.table.DefaultTableModel;
import java.awt.*;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.math.BigDecimal;
import java.nio.file.Path;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.List;

public class ZapisniciGuiApp extends JFrame {
    
    private static final Path RASKRSNICE_XLSX = Path.of("src/main/resources/Nazivi_raskrsnica.xlsx");
    private static final Path CENOVNIK_XLSX   = Path.of("src/main/resources/Blanko za izradu zapisnika.xlsx");
    private static final Path TEMPLATE_DOCX   = Path.of("src/main/resources/template.docx");
    private static final Path OUTPUT_DIR      = Path.of("output");
    
    private final RaskrsniceRepository raskrsniceRepository;
    private final CenovnikLookup cenovnikLookup;
    private final ExcelReaderService excelReaderService;
    private final ZapisnikService zapisnikService;
    
    private final JTextField datumField = new JTextField(12);
    private final JTextField kBrojField = new JTextField(12);
    private final JTextField brojField = new JTextField(12);
    private final JLabel raskrsnicaLabel = new JLabel("Naziv raskrsnice: ");
    
    private final DefaultTableModel tableModel = new DefaultTableModel(
            new Object[]{"R.br", "Količina"}, 0
    );
    
    private final JTable stavkeTable = new JTable(tableModel);
    
    private final JLabel ukupanZbirProgramaLabel = new JLabel("Ukupan zbir programa: 0.00");
    private BigDecimal ukupanZbirPrograma = BigDecimal.ZERO;
    
    public ZapisniciGuiApp() throws Exception {
        this.raskrsniceRepository = new RaskrsniceRepository(RASKRSNICE_XLSX);
        this.cenovnikLookup = new ExcelCenovnik(CENOVNIK_XLSX);
        this.excelReaderService = new ExcelReaderService(cenovnikLookup, raskrsniceRepository);
        this.zapisnikService = new ZapisnikService(excelReaderService);
        
        initUi();
    }
    
    private void initUi() {
        setTitle("Zapisnici");
        setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        setSize(700, 500);
        setLocationRelativeTo(null);
        
        JPanel mainPanel = new JPanel(new BorderLayout(10, 10));
        mainPanel.setBorder(BorderFactory.createEmptyBorder(10, 10, 10, 10));
        
        JPanel topPanel = new JPanel(new GridBagLayout());
        GridBagConstraints gbc = new GridBagConstraints();
        gbc.insets = new Insets(5, 5, 5, 5);
        gbc.anchor = GridBagConstraints.WEST;
        
        datumField.setText(LocalDate.now().format(DateTimeFormatter.ofPattern("dd.MM.yyyy.")));
        
        gbc.gridx = 0;
        gbc.gridy = 0;
        topPanel.add(new JLabel("Datum (dd.MM.yyyy.):"), gbc);
        
        gbc.gridx = 1;
        topPanel.add(datumField, gbc);
        
        gbc.gridx = 0;
        gbc.gridy = 1;
        topPanel.add(new JLabel("K broj:"), gbc);
        
        gbc.gridx = 1;
        topPanel.add(kBrojField, gbc);
        
        gbc.gridx = 0;
        gbc.gridy = 2;
        topPanel.add(new JLabel("Broj zapisnika:"), gbc);
        
        gbc.gridx = 1;
        topPanel.add(brojField, gbc);
        
        gbc.gridx = 0;
        gbc.gridy = 3;
        gbc.gridwidth = 2;
        topPanel.add(raskrsnicaLabel, gbc);
        
        JButton proveriRaskrsnicuBtn = new JButton("Pronađi raskrsnicu");
        proveriRaskrsnicuBtn.addActionListener(e -> prikaziNazivRaskrsnice());
        
        gbc.gridy = 4;
        gbc.gridwidth = 1;
        topPanel.add(proveriRaskrsnicuBtn, gbc);
        
        mainPanel.add(topPanel, BorderLayout.NORTH);
        
        JScrollPane scrollPane = new JScrollPane(stavkeTable);
        mainPanel.add(scrollPane, BorderLayout.CENTER);
        
        JPanel buttonsPanel = new JPanel(new FlowLayout(FlowLayout.LEFT));
        
        JButton addRowBtn = new JButton("Dodaj stavku");
        addRowBtn.addActionListener(e -> tableModel.addRow(new Object[]{"", ""}));
        
        JButton removeRowBtn = new JButton("Obriši stavku");
        removeRowBtn.addActionListener(e -> obrisiSelektovanuStavku());
        
        JButton generisiBtn = new JButton("Generiši zapisnik");
        generisiBtn.addActionListener(e -> generisiZapisnik(false));
        
        JButton generisiISstampaBtn = new JButton("Generiši i štampaj");
        generisiISstampaBtn.addActionListener(e -> generisiZapisnik(true));
        
        JButton noviUnosBtn = new JButton("Novi unos");
        noviUnosBtn.addActionListener(e -> resetForm());
        
        buttonsPanel.add(addRowBtn);
        buttonsPanel.add(removeRowBtn);
        buttonsPanel.add(generisiBtn);
        buttonsPanel.add(generisiISstampaBtn);
        buttonsPanel.add(noviUnosBtn);
        
        JPanel bottomPanel = new JPanel(new BorderLayout());
        bottomPanel.add(buttonsPanel, BorderLayout.NORTH);
        bottomPanel.add(ukupanZbirProgramaLabel, BorderLayout.SOUTH);
        
        mainPanel.add(bottomPanel, BorderLayout.SOUTH);
        
        setContentPane(mainPanel);
    }
    
    private void prikaziNazivRaskrsnice() {
        try {
            String kBroj = normalizeKBroj(kBrojField.getText());
            String nazivRaskrsnice = raskrsniceRepository.getNazivRaskrsnice(kBroj);
            
            if (nazivRaskrsnice == null || nazivRaskrsnice.isBlank()) {
                raskrsnicaLabel.setText("Naziv raskrsnice: nije pronađena za " + kBroj);
                return;
            }
            
            raskrsnicaLabel.setText("Naziv raskrsnice: " + nazivRaskrsnice);
        } catch (Exception ex) {
            showError(ex.getMessage());
        }
    }
    
    private void obrisiSelektovanuStavku() {
        int selectedRow = stavkeTable.getSelectedRow();
        if (selectedRow >= 0) {
            tableModel.removeRow(selectedRow);
        } else {
            showError("Selektuj red koji želiš da obrišeš.");
        }
    }
    
    private void generisiZapisnik(boolean stampajPosle) {
        try {
            LocalDate datum = LocalDate.parse(
                    datumField.getText().trim(),
                    DateTimeFormatter.ofPattern("dd.MM.yyyy.")
            );
            
            String kBroj = normalizeKBroj(kBrojField.getText());
            String broj = brojField.getText().trim();
            
            if (broj.isBlank()) {
                throw new IllegalArgumentException("Broj zapisnika ne sme biti prazan.");
            }
            
            String nazivRaskrsnice = raskrsniceRepository.getNazivRaskrsnice(kBroj);
            if (nazivRaskrsnice == null || nazivRaskrsnice.isBlank()) {
                throw new IllegalArgumentException("Ne postoji raskrsnica za K broj: " + kBroj);
            }
            
            List<Integer> redniBrojevi = new ArrayList<>();
            List<BigDecimal> kolicine = new ArrayList<>();
            
            for (int i = 0; i < tableModel.getRowCount(); i++) {
                Object rbObj = tableModel.getValueAt(i, 0);
                Object kolObj = tableModel.getValueAt(i, 1);
                
                String rbText = rbObj == null ? "" : rbObj.toString().trim();
                String kolText = kolObj == null ? "" : kolObj.toString().trim();
                
                if (rbText.isEmpty() && kolText.isEmpty()) {
                    continue;
                }
                
                if (rbText.isEmpty() || kolText.isEmpty()) {
                    throw new IllegalArgumentException("Svaki red mora imati i R.br i količinu.");
                }
                
                int rb = Integer.parseInt(rbText);
                BigDecimal kolicina = new BigDecimal(kolText.replace(",", "."));
                
                redniBrojevi.add(rb);
                kolicine.add(kolicina);
            }
            
            if (redniBrojevi.isEmpty()) {
                throw new IllegalArgumentException("Moraš uneti bar jednu stavku.");
            }
            
            ZapisnikMetadata metadata = new ZapisnikMetadata(datum, kBroj, broj, nazivRaskrsnice);
            
            String outName = buildOutputFileName(metadata);
            Path outPath = OUTPUT_DIR.resolve(outName);
            
            BigDecimal zbirUkupno;
            try (FileInputStream in = new FileInputStream(TEMPLATE_DOCX.toFile());
                 XWPFDocument doc = new XWPFDocument(in)) {
                
                zbirUkupno = zapisnikService.generisiZapisnik(doc, metadata, redniBrojevi, kolicine);
                
                OUTPUT_DIR.toFile().mkdirs();
                try (FileOutputStream out = new FileOutputStream(outPath.toFile())) {
                    doc.write(out);
                }
            }
            
            ukupanZbirPrograma = ukupanZbirPrograma.add(zbirUkupno);
            ukupanZbirProgramaLabel.setText(
                    "Ukupan zbir programa: " + WordTabelaPopunjavanje.formatMoney(ukupanZbirPrograma)
            );
            
            if (stampajPosle) {
                printWordDocument(outPath);
            }
            
            JOptionPane.showMessageDialog(
                    this,
                    "Zapisnik je uspešno sačuvan:\n" + outPath.toAbsolutePath(),
                    "Uspeh",
                    JOptionPane.INFORMATION_MESSAGE
            );
            
        } catch (Exception ex) {
            showError(ex.getMessage());
        }
    }
    
    private void resetForm() {
        datumField.setText(LocalDate.now().format(DateTimeFormatter.ofPattern("dd.MM.yyyy.")));
        kBrojField.setText("");
        brojField.setText("");
        raskrsnicaLabel.setText("Naziv raskrsnice: ");
        tableModel.setRowCount(0);
    }
    
    private String normalizeKBroj(String input) {
        String value = input == null ? "" : input.trim().toUpperCase();
        if (value.isEmpty()) {
            throw new IllegalArgumentException("K broj ne sme biti prazan.");
        }
        if (!value.startsWith("K")) {
            value = "K" + value;
        }
        return value;
    }
    
    private String buildOutputFileName(ZapisnikMetadata m) {
        return m.getBroj() + ".docx";
    }
    
    private void printWordDocument(Path documentPath) throws Exception {
        if (documentPath == null) {
            throw new IllegalArgumentException("Putanja do dokumenta ne sme biti null.");
        }
        
        if (!Desktop.isDesktopSupported()) {
            throw new UnsupportedOperationException("Desktop API nije podržan na ovom sistemu.");
        }
        
        Desktop desktop = Desktop.getDesktop();
        
        if (!desktop.isSupported(Desktop.Action.PRINT)) {
            throw new UnsupportedOperationException("PRINT akcija nije podržana na ovom sistemu.");
        }
        
        File file = documentPath.toFile();
        if (!file.exists()) {
            throw new IllegalStateException("Dokument za štampu ne postoji: " + documentPath);
        }
        
        desktop.print(file);
    }
    
    private void showError(String message) {
        JOptionPane.showMessageDialog(
                this,
                message,
                "Greška",
                JOptionPane.ERROR_MESSAGE
        );
    }
    
    public static void main(String[] args) {
        SwingUtilities.invokeLater(() -> {
            try {
                new ZapisniciGuiApp().setVisible(true);
            } catch (Exception e) {
                JOptionPane.showMessageDialog(
                        null,
                        "Pokretanje aplikacije nije uspelo:\n" + e.getMessage(),
                        "Greška",
                        JOptionPane.ERROR_MESSAGE
                );
            }
        });
    }
}

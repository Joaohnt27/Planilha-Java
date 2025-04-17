import javax.imageio.ImageIO;
import javax.swing.*;
import java.awt.*;
import java.awt.Font;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.time.*;
import java.time.format.TextStyle;
import java.util.*;
import java.util.List;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class FormularioEstagiario extends JFrame {

    private JTextField campoNome;
    private JTextField campoCodigo;
    private JComboBox comboSetor;
    private JComboBox comboPeriodo;
    private JComboBox<String> comboPagamento;
    private JButton btnGerarPlanilha;

    public FormularioEstagiario() {
        setTitle("Gerar Folha de Frequência");
        setSize(500, 500);
        setLocationRelativeTo(null);
        setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        setLayout(new BorderLayout());

        try {
            InputStream imgUnaerp = getClass().getClassLoader().getResourceAsStream("Logo.png");
            if (imgUnaerp != null) {
                ImageIcon iconFrequencia = new ImageIcon(ImageIO.read(imgUnaerp));
                setIconImage(iconFrequencia.getImage());
            } else {
                System.err.println("A imagem não foi encontrada no classpath!");
            }
        } catch (Exception e) {
            e.printStackTrace();
        }

        // Painel principal vertical
        JPanel mainPanel = new JPanel();
        mainPanel.setLayout(new BoxLayout(mainPanel, BoxLayout.Y_AXIS));

        // Painel da logo
        JLabel imagemLabel = new JLabel();
        try {
            ImageIcon imagem = new ImageIcon(getClass().getClassLoader().getResource("Logo.png"));
            imagemLabel.setIcon(imagem);
            imagemLabel.setAlignmentX(Component.CENTER_ALIGNMENT);
        } catch (Exception e) {
            e.printStackTrace();
        }

        mainPanel.add(Box.createVerticalStrut(10));
        mainPanel.add(imagemLabel);

        // Painel de campos
        JPanel painelCampos = new JPanel(new GridLayout(8, 2, 5, 5));
        campoNome = new JTextField();
        campoCodigo = new JTextField();
        comboSetor = new JComboBox<>(new String[] { "CIT - Service Desk", "CIT - Desenvolvimento", "LIAPE" });
        comboPeriodo = new JComboBox<>(new String[] {"Manhã", "Tarde", "Noite"});
        comboPagamento = new JComboBox<>(new String[] {"100% Bolsa", "R$ 1.400,00"});
        btnGerarPlanilha = new JButton("Gerar Planilha");

        painelCampos.add(new JLabel("Nome:"));
        painelCampos.add(campoNome);
        painelCampos.add(new JLabel("Código:"));
        painelCampos.add(campoCodigo);
        painelCampos.add(new JLabel("Setor:"));
        painelCampos.add(comboSetor);
        painelCampos.add(new JLabel("Período:"));
        painelCampos.add(comboPeriodo);
        painelCampos.add(new JLabel("Bolsa/Salário:"));
        painelCampos.add(comboPagamento);
        painelCampos.add(new JLabel(""));
        painelCampos.add(btnGerarPlanilha);

        mainPanel.add(Box.createVerticalStrut(20));
        mainPanel.add(painelCampos);

        // Rodapé
        JPanel painelInferior = new JPanel();
        painelInferior.setLayout(new BoxLayout(painelInferior, BoxLayout.Y_AXIS));
        JLabel rodape = new JLabel("@2025 | Versão: 1.0 - Desenvolvido por João Henrique Nazar Tavares");
        rodape.setFont(new Font("Calibri", Font.BOLD, 10));
        rodape.setAlignmentX(Component.CENTER_ALIGNMENT);
        painelInferior.add(Box.createVerticalStrut(10));
        painelInferior.add(rodape);

        add(mainPanel, BorderLayout.CENTER);
        add(painelInferior, BorderLayout.SOUTH);

        btnGerarPlanilha.addActionListener(e -> gerarPlanilha());

        setVisible(true);
    }

    private void gerarPlanilha() {
        String nome = campoNome.getText();
        String codigo = campoCodigo.getText();
        String setor = (String) comboSetor.getSelectedItem();
        String periodo = (String) comboPeriodo.getSelectedItem();
        String bolsa = (String)comboPagamento.getSelectedItem();

        try {
            InputStream fis = getClass().getClassLoader().getResourceAsStream("frequencia2025.xlsx");
            Workbook workbook = new XSSFWorkbook(fis);
            Sheet sheet = workbook.getSheetAt(0);

            setCellValueSafe(sheet, 5, 2, nome);     // C6
            setCellValueSafe(sheet, 6, 2, codigo);   // C7
            setCellValueSafe(sheet, 7, 3, setor);    // D8
            setCellValueSafe(sheet, 10, 4, bolsa);   // E11

            String mesAtual = LocalDate.now().getMonth().getDisplayName(TextStyle.FULL, new Locale("pt", "BR")).toUpperCase();
            setCellValueSafe(sheet, 2, 4, "MÊS DE " + mesAtual);  // Ex: célula E3

            List<LocalDate> dias = gerarDiasEstagio();

            String entradaPadrao = "07:00", saidaPadrao = "13:00";
            if ("Tarde".equals(periodo)) {
                entradaPadrao = "13:00"; saidaPadrao = "19:00";
            } else if ("Noite".equals(periodo)) {
                entradaPadrao = "16:00"; saidaPadrao = "22:00";
            }

            CellStyle estiloMes = sheet.getRow(14).getCell(0).getCellStyle();
            CellStyle estiloDia = sheet.getRow(14).getCell(1).getCellStyle();
            CellStyle estiloSemana = sheet.getRow(14).getCell(2).getCellStyle();
            CellStyle estiloEntrada = sheet.getRow(14).getCell(3).getCellStyle();
            CellStyle estiloSaida = sheet.getRow(14).getCell(4).getCellStyle();

            int linhaInicio = 14;
            for (int i = 0; i < dias.size(); i++) {
                LocalDate data = dias.get(i);
                Row row = sheet.getRow(linhaInicio + i);
                if (row == null) row = sheet.createRow(linhaInicio + i);

                String nomeMes = data.getMonth().getDisplayName(TextStyle.FULL, new Locale("pt", "BR")).toUpperCase();
                String diaDoMes = String.valueOf(data.getDayOfMonth());
                String diaDaSemana = data.getDayOfWeek().getDisplayName(TextStyle.FULL, new Locale("pt", "BR"));

                String entrada = "-";
                String saida = "-";
                if (!(data.getDayOfWeek() == DayOfWeek.SATURDAY || data.getDayOfWeek() == DayOfWeek.SUNDAY)) {
                    entrada = entradaPadrao;
                    saida = saidaPadrao;
                }

                Cell cellMes = row.createCell(0);
                cellMes.setCellValue(nomeMes);
                cellMes.setCellStyle(estiloMes);

                Cell cellDia = row.createCell(1);
                cellDia.setCellValue(diaDoMes);
                cellDia.setCellStyle(estiloDia);

                Cell cellSemana = row.createCell(2);
                cellSemana.setCellValue(diaDaSemana);
                cellSemana.setCellStyle(estiloSemana);

                Cell cellEntrada = row.createCell(3);
                cellEntrada.setCellValue(entrada);
                cellEntrada.setCellStyle(estiloEntrada);

                Cell cellSaida = row.createCell(4);
                cellSaida.setCellValue(saida);
                cellSaida.setCellStyle(estiloSaida);
            }

            String nomeArquivo = System.getProperty("user.home") + "/Downloads/frequencia_" + nome.replace(" ", "_") + ".xlsx";
            FileOutputStream fos = new FileOutputStream(nomeArquivo);
            workbook.write(fos);
            fos.close();
            workbook.close();

            JOptionPane.showMessageDialog(this, "Planilha gerada com sucesso!");
        } catch (Exception ex) {
            ex.printStackTrace();
            JOptionPane.showMessageDialog(this, "Erro ao gerar planilha!");
        }
    }

    private void setCellValueSafe(Sheet sheet, int rowIndex, int colIndex, String value) {
        Row row = sheet.getRow(rowIndex);
        if (row == null) row = sheet.createRow(rowIndex);
        Cell cell = row.getCell(colIndex);
        if (cell == null) cell = row.createCell(colIndex);
        cell.setCellValue(value);
    }

    private List<LocalDate> gerarDiasEstagio() {
        LocalDate hoje = LocalDate.now();
        LocalDate inicio = LocalDate.of(
                hoje.getMonthValue() == 1 ? hoje.getYear() - 1 : hoje.getYear(),
                hoje.getMonthValue() == 1 ? 12 : hoje.getMonthValue() - 1,
                21
        );
        LocalDate fim = LocalDate.of(hoje.getYear(), hoje.getMonth(), 20);

        List<LocalDate> dias = new ArrayList<>();
        while (!inicio.isAfter(fim)) {
            dias.add(inicio);
            inicio = inicio.plusDays(1);
        }
        return dias;
    }

    public static void main(String[] args) {
        SwingUtilities.invokeLater(FormularioEstagiario::new);
    }
}

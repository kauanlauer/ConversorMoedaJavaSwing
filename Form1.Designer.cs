import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.*;
import javax.swing.table.DefaultTableModel;
import java.awt.*;
import java.awt.event.*;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

public class ManipuladorDePedidosSwing extends JFrame
{

    private JTable table;
private DefaultTableModel tableModel;

public ManipuladorDePedidosSwing()
{
    setTitle("Manipulador de Pedidos");
    setSize(800, 600);
    setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
    setLocationRelativeTo(null);

    // Configuração da tabela
    tableModel = new DefaultTableModel();
    table = new JTable(tableModel);
    table.setComponentPopupMenu(createPopupMenu());

    // Botão para abrir a planilha
    JButton btnAbrir = new JButton("Abrir Planilha");
    btnAbrir.addActionListener(e->abrirPlanilha());

    // Layout
    setLayout(new BorderLayout());
    add(new JScrollPane(table), BorderLayout.CENTER);
    add(btnAbrir, BorderLayout.SOUTH);
}

private void abrirPlanilha()
{
    JFileChooser fileChooser = new JFileChooser();
    int result = fileChooser.showOpenDialog(this);

    if (result == JFileChooser.APPROVE_OPTION)
    {
        File file = fileChooser.getSelectedFile();
        try (FileInputStream fis = new FileInputStream(file);
        Workbook workbook = new XSSFWorkbook(fis)) {

            // Abrir a guia "Pendentes"
            Sheet sheet = workbook.getSheet("Pendentes");
            if (sheet != null)
            {
                tableModel.setRowCount(0);
                tableModel.setColumnCount(sheet.getRow(0).getPhysicalNumberOfCells());

                for (Row row : sheet) {
                        Object[] rowData = new Object[row.getPhysicalNumberOfCells()];
for (int i = 0; i < row.getPhysicalNumberOfCells(); i++)
{
    Cell cell = row.getCell(i);
    rowData[i] = getCellValue(cell);
}
tableModel.addRow(rowData);
                    }
                } else
{
    JOptionPane.showMessageDialog(this, "A guia 'Pendentes' não foi encontrada!");
}
            } catch (IOException ex) {
                JOptionPane.showMessageDialog(this, "Erro ao abrir o arquivo: " + ex.getMessage());
            }
        }
    }

    private Object getCellValue(Cell cell)
{
    switch (cell.getCellType())
    {
        case STRING:
            return cell.getStringCellValue();
        case NUMERIC:
            if (DateUtil.isCellDateFormatted(cell))
            {
                return cell.getDateCellValue();
            }
            else
            {
                return cell.getNumericCellValue();
            }
        case BOOLEAN:
            return cell.getBooleanCellValue();
        default:
            return "";
    }
}

private JPopupMenu createPopupMenu()
{
    JPopupMenu popupMenu = new JPopupMenu();

    JMenuItem deleteItem = new JMenuItem("Excluir Linha");
    deleteItem.addActionListener(e->excluirLinha());

    JMenuItem copyItem = new JMenuItem("Copiar Linha");
    copyItem.addActionListener(e->copiarLinha());

    JMenuItem duplicateItem = new JMenuItem("Duplicar Linha");
    duplicateItem.addActionListener(e->duplicarLinha());

    popupMenu.add(deleteItem);
    popupMenu.add(copyItem);
    popupMenu.add(duplicateItem);

    return popupMenu;
}

private void excluirLinha()
{
    int selectedRow = table.getSelectedRow();
    if (selectedRow != -1)
    {
        tableModel.removeRow(selectedRow);
    }
}

private void copiarLinha()
{
    int selectedRow = table.getSelectedRow();
    if (selectedRow != -1)
    {
        Object[] rowData = new Object[tableModel.getColumnCount()];
        for (int i = 0; i < tableModel.getColumnCount(); i++)
        {
            rowData[i] = tableModel.getValueAt(selectedRow, i);
        }
        // Copia os dados da linha selecionada para o clipboard ou manipula conforme necessário
        JOptionPane.showMessageDialog(this, "Linha copiada: " + java.util.Arrays.toString(rowData));
    }
}

private void duplicarLinha()
{
    int selectedRow = table.getSelectedRow();
    if (selectedRow != -1)
    {
        Object[] rowData = new Object[tableModel.getColumnCount()];
        for (int i = 0; i < tableModel.getColumnCount(); i++)
        {
            rowData[i] = tableModel.getValueAt(selectedRow, i);
        }
        tableModel.addRow(rowData);
    }
}

public static void main(String[] args)
{
    SwingUtilities.invokeLater(()-> {
        ManipuladorDePedidosSwing frame = new ManipuladorDePedidosSwing();
        frame.setVisible(true);
    });
}
}

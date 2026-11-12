import java.awt.Color;
import java.awt.Component;
import java.awt.Font;
import javax.swing.BorderFactory;
import javax.swing.JLabel;
import javax.swing.JList;
import javax.swing.JTable;
import javax.swing.ListCellRenderer;
import javax.swing.SwingConstants;
import javax.swing.table.DefaultTableCellRenderer;

/**
 * Rendering layer: paints each cell according to its {@link CellStyle}
 * (bold / italic / size, alignment, fill and text colour) with subtle grid
 * lines, a tinted selection and a green "active cell" frame like Excel.
 */
public class SpreadsheetCellRenderer extends DefaultTableCellRenderer {

    public static final Color GREEN = new Color(0x107C41);
    static final Color GRID = new Color(0xD9D9D9);
    static final Color HEADER_BG = new Color(0xF3F3F3);
    static final Color HEADER_SELECTED_BG = new Color(0xC6E5D1);
    static final Color HEADER_HOVER_BG = new Color(0xE3F1E8);
    private static final Color ERROR_FG = new Color(0xC00000);

    private final SpreadsheetModel model;
    private double zoom = 1.0;
    /** Cell under the mouse (-1 = none); gets a faint highlight. Set by the sheet view. */
    public int hoverRow = -1;
    public int hoverCol = -1;

    public SpreadsheetCellRenderer(SpreadsheetModel model) {
        this.model = model;
    }

    /** Scale factor for explicitly sized fonts (the table's own font is already scaled). */
    public void setZoom(double zoom) {
        this.zoom = zoom;
    }

    @Override
    public Component getTableCellRendererComponent(JTable table, Object value, boolean isSelected,
                                                   boolean hasFocus, int row, int column) {
        JLabel label = (JLabel) super.getTableCellRendererComponent(table, value, false, false, row, column);
        CellStyle st = model.getStyle(row, column);

        // Font
        Font base = table.getFont();
        int flags = Font.PLAIN;
        if (st != null && st.bold) flags |= Font.BOLD;
        if (st != null && st.italic) flags |= Font.ITALIC;
        float size = (st != null && st.fontSize > 0) ? (float) (st.fontSize * zoom) : base.getSize2D();
        label.setFont(base.deriveFont(flags, size));

        // Alignment: explicit style wins; otherwise numbers right, errors centred, text left
        int align;
        if (st != null && st.alignment != CellStyle.ALIGN_AUTO) align = st.alignment;
        else if (model.isFormulaError(row, column) || model.getComputedValue(row, column) instanceof Boolean) {
            align = SwingConstants.CENTER; // errors and TRUE/FALSE are centred, like Excel
        }
        else if (model.isNumeric(row, column)) align = SwingConstants.RIGHT;
        else align = SwingConstants.LEFT;
        label.setHorizontalAlignment(align);

        // Colours
        Color bg = (st != null && st.background != null) ? st.background : Color.WHITE;
        Color fg = (st != null && st.foreground != null) ? st.foreground : Color.BLACK;
        if (model.isFormulaError(row, column) && (st == null || st.foreground == null)) fg = ERROR_FG;

        boolean active = row == table.getSelectionModel().getLeadSelectionIndex()
                && column == table.getColumnModel().getSelectionModel().getLeadSelectionIndex();
        if (isSelected && !active) bg = blend(bg, GREEN, 0.20f); // active cell stays "clean", like Excel
        else if (!isSelected && row == hoverRow && column == hoverCol) bg = blend(bg, GREEN, 0.07f);
        label.setBackground(bg);
        label.setForeground(fg);
        label.setOpaque(true);

        // Borders
        if (active) {
            label.setBorder(BorderFactory.createCompoundBorder(
                    BorderFactory.createLineBorder(GREEN, 2),
                    BorderFactory.createEmptyBorder(0, 2, 0, 2)));
        } else {
            label.setBorder(BorderFactory.createCompoundBorder(
                    BorderFactory.createMatteBorder(0, 0, 1, 1, GRID),
                    BorderFactory.createEmptyBorder(0, 4, 0, 4)));
        }
        return label;
    }

    private static Color blend(Color a, Color b, float amount) {
        float inv = 1f - amount;
        return new Color(
                Math.round(a.getRed() * inv + b.getRed() * amount),
                Math.round(a.getGreen() * inv + b.getGreen() * amount),
                Math.round(a.getBlue() * inv + b.getBlue() * amount));
    }

    // ------------------------------------------------------------------

    /** Column header (A, B, C ...) - highlights columns that contain the selection. */
    public static class ColumnHeaderRenderer extends DefaultTableCellRenderer {
        public int hover = -1; // header under the mouse

        @Override
        public Component getTableCellRendererComponent(JTable table, Object value, boolean isSelected,
                                                       boolean hasFocus, int row, int column) {
            JLabel l = (JLabel) super.getTableCellRendererComponent(table, value, false, false, row, column);
            boolean sel = table.isColumnSelected(column);
            l.setHorizontalAlignment(SwingConstants.CENTER);
            l.setFont(table.getFont().deriveFont(Font.BOLD, Math.max(7f, table.getFont().getSize2D() - 1f)));
            l.setOpaque(true);
            l.setBackground(sel ? HEADER_SELECTED_BG : column == hover ? HEADER_HOVER_BG : HEADER_BG);
            l.setForeground(sel || column == hover ? new Color(0x0B5D31) : new Color(0x444444));
            l.setBorder(BorderFactory.createMatteBorder(0, 0, 1, 1, new Color(0xBFBFBF)));
            return l;
        }
    }

    /** Row header (1, 2, 3 ...) - used as the JList in the scroll pane's row header. */
    public static class RowHeaderRenderer extends JLabel implements ListCellRenderer<String> {
        private final JTable table;
        public int hover = -1; // row header under the mouse

        public RowHeaderRenderer(JTable table) {
            this.table = table;
            setOpaque(true);
            setHorizontalAlignment(SwingConstants.CENTER);
        }

        @Override
        public Component getListCellRendererComponent(JList<? extends String> list, String value, int index,
                                                      boolean isSelected, boolean cellHasFocus) {
            boolean sel = table.isRowSelected(index);
            setText(value);
            setFont(table.getFont().deriveFont(Font.BOLD, Math.max(7f, table.getFont().getSize2D() - 1f)));
            setBackground(sel ? HEADER_SELECTED_BG : index == hover ? HEADER_HOVER_BG : HEADER_BG);
            setForeground(sel || index == hover ? new Color(0x0B5D31) : new Color(0x444444));
            setBorder(BorderFactory.createMatteBorder(0, 0, 1, 1, new Color(0xBFBFBF)));
            return this;
        }
    }
}

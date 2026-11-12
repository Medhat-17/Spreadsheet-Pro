import java.awt.BorderLayout;
import java.awt.CardLayout;
import java.awt.GridBagLayout;
import java.util.EnumMap;
import java.util.Map;
import javax.swing.BoxLayout;
import javax.swing.JCheckBox;
import javax.swing.JSlider;
import javax.swing.SwingUtilities;
import java.awt.FontMetrics;
import java.awt.Point;
import java.awt.event.MouseWheelEvent;
import java.awt.event.MouseWheelListener;
import javax.swing.event.TableColumnModelEvent;
import javax.swing.event.TableColumnModelListener;
import javax.swing.table.JTableHeader;
import javax.swing.table.TableColumn;
import java.awt.Color;
import java.awt.Component;
import java.awt.Cursor;
import java.awt.BasicStroke;
import java.awt.Rectangle;
import java.awt.Dimension;
import java.awt.GridLayout;
import java.awt.Toolkit;
import java.awt.datatransfer.Clipboard;
import java.awt.datatransfer.DataFlavor;
import java.awt.event.WindowAdapter;
import java.awt.event.WindowEvent;
import java.awt.image.BufferedImage;
import java.util.Arrays;
import javax.imageio.ImageIO;
import javax.swing.Icon;
import javax.swing.JPasswordField;
import javax.swing.event.AncestorEvent;
import javax.swing.event.AncestorListener;
import javax.swing.filechooser.FileNameExtensionFilter;
import java.awt.FlowLayout;
import java.awt.Font;
import java.awt.Graphics;
import java.awt.Graphics2D;
import java.awt.RenderingHints;
import java.awt.event.ActionEvent;
import java.awt.event.InputEvent;
import java.awt.event.KeyEvent;
import java.awt.event.MouseAdapter;
import java.awt.event.MouseEvent;
import java.io.File;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.util.ArrayList;
import java.util.List;
import java.util.function.Consumer;
import javax.swing.AbstractAction;
import javax.swing.BorderFactory;
import javax.swing.ButtonModel;
import javax.swing.DefaultCellEditor;
import javax.swing.JButton;
import javax.swing.JComboBox;
import javax.swing.JComponent;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JList;
import javax.swing.JMenu;
import javax.swing.JMenuBar;
import javax.swing.JMenuItem;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.JPopupMenu;
import javax.swing.JScrollPane;
import javax.swing.JTabbedPane;
import javax.swing.JTable;
import javax.swing.JTextArea;
import javax.swing.JTextField;
import javax.swing.KeyStroke;
import javax.swing.ListSelectionModel;
import javax.swing.SwingConstants;
import javax.swing.SwingUtilities;
import javax.swing.UIManager;
import javax.swing.border.CompoundBorder;
import javax.swing.border.EmptyBorder;
import javax.swing.border.MatteBorder;
import javax.swing.event.ListSelectionListener;
import javax.swing.plaf.basic.BasicButtonUI;

/**
 * UI &amp; layout layer: the main window.
 *
 * <ul>
 *   <li>Excel-style green ribbon (#107C41) with all tools</li>
 *   <li>Formula bar: cell badge, "fx" label and a formula text field synced with the selection</li>
 *   <li>Multi-sheet workbook in a JTabbedPane (add / rename / delete tabs)</li>
 *   <li>Status bar with Count / Sum / Average of the selected cells</li>
 * </ul>
 * Each worksheet is a {@link SpreadsheetModel} shown in a JTable that uses
 * {@link SpreadsheetCellRenderer}; formulas are evaluated by {@link FormulaEngine}.
 */
public class SpreadsheetApp extends JFrame {

    static final Color GREEN = new Color(0x107C41);
    static final Color LIGHT_GRAY = new Color(0xF3F3F3);
    static final Color GRID = new Color(0xD0D0D0);

    private final JTabbedPane tabs = new JTabbedPane();
    private final JLabel cellBadge = new JLabel("A1", SwingConstants.CENTER);
    private final JTextField formulaField = new JTextField();
    private final JLabel statusLabel = new JLabel("Ready");
    private final JLabel statsLabel = new JLabel("");
    private final JComboBox<Integer> sizeCombo = new JComboBox<>(new Integer[]{9, 10, 11, 12, 13, 14, 16, 18, 20});
    private final JPanel fillChip = new JPanel();

    private Color currentBackground = new Color(0xFFF2CC);
    private int sheetCounter = 0;
    private boolean updatingUi = false;
    private final JLabel zoomLabel = new JLabel("100%");
    private final JSlider zoomSlider = new JSlider(SpreadsheetModel.MIN_ZOOM, SpreadsheetModel.MAX_ZOOM, 100);
    private MacroIDEDialog macroDialog;

    private final JLabel protLabel = new JLabel("", SwingConstants.CENTER);
    private RibbonButton protectButton;
    private static final Color RIBBON_DARK = new Color(0x0B5D31);
    private final CardLayout ribbonCards = new CardLayout();
    private JPanel ribbonBody;
    private String ribbonSelected = "Home";
    private final List<RibbonTab> ribbonTabs = new ArrayList<>();
    private RibbonButton boldBtn, italicBtn, alignLeftBtn, alignCenterBtn, alignRightBtn;
    private final DrawingLayer.Settings draw = new DrawingLayer.Settings();
    private final Map<DrawingLayer.Tool, RibbonButton> toolButtons = new EnumMap<>(DrawingLayer.Tool.class);
    private final List<SwatchButton> swatches = new ArrayList<>();
    private final JPanel drawChip = new JPanel();
    private final JLabel widthLabel = new JLabel("3 px");
    private File currentFile;
    private char[] workbookPassword;
    private boolean dirty = false;
    private boolean loading = false;

    public static void main(String[] args) {
        SwingUtilities.invokeLater(() -> {
            try {
                UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName());
            } catch (Exception ignored) {
                // fall back to the default look and feel
            }
            new SpreadsheetApp().setVisible(true);
        });
    }

    public SpreadsheetApp() {
        super("Spreadsheet Pro");
        setDefaultCloseOperation(DO_NOTHING_ON_CLOSE);
        addWindowListener(new WindowAdapter() {
            @Override
            public void windowClosing(WindowEvent e) {
                exitApp();
            }
        });
        setIconImage(Icons.logo(64));
        setSize(1360, 780);
        setLocationRelativeTo(null);
        setJMenuBar(buildMenuBar());

        JPanel north = new JPanel(new BorderLayout());
        north.add(buildRibbon(), BorderLayout.NORTH);
        north.add(buildFormulaBar(), BorderLayout.SOUTH);
        add(north, BorderLayout.NORTH);

        tabs.setTabPlacement(JTabbedPane.BOTTOM);
        tabs.addChangeListener(e -> syncFormulaBar());
        add(tabs, BorderLayout.CENTER);
        add(buildStatusBar(), BorderLayout.SOUTH);

        addSheet(nextSheetName());
        dirty = false;
        updateTitle();
    }

    // ==================================================================
    // Ribbon
    // ==================================================================

    private JComponent buildRibbon() {
        JPanel root = new JPanel(new BorderLayout());
        root.setBackground(GREEN);

        JPanel strip = new JPanel(new FlowLayout(FlowLayout.LEFT, 2, 0));
        strip.setBackground(RIBBON_DARK);
        strip.setBorder(new EmptyBorder(4, 8, 0, 8));

        ribbonBody = new JPanel(ribbonCards);
        ribbonBody.setBackground(GREEN);
        ribbonBody.add(homeTab(), "Home");
        ribbonBody.add(insertTab(), "Insert");
        ribbonBody.add(drawTab(), "Draw");
        ribbonBody.add(viewTab(), "View");
        for (String name : new String[]{"Home", "Insert", "Draw", "View"}) {
            RibbonTab t = new RibbonTab(name);
            ribbonTabs.add(t);
            strip.add(t);
        }
        root.add(strip, BorderLayout.NORTH);
        root.add(ribbonBody, BorderLayout.CENTER);
        selectRibbonTab("Home");
        return root;
    }

    private void selectRibbonTab(String name) {
        ribbonSelected = name;
        ribbonCards.show(ribbonBody, name);
        for (RibbonTab t : ribbonTabs) t.repaint();
        if (!"Draw".equals(name)) setDrawTool(DrawingLayer.Tool.NONE); // leaving Draw switches drawing off
    }

    private JPanel tabPanel() {
        JPanel p = new JPanel(new FlowLayout(FlowLayout.LEFT, 0, 4));
        p.setBackground(GREEN);
        return p;
    }

    /** A titled cluster of ribbon controls with a thin divider on its right. */
    private JComponent group(String title, JComponent... items) {
        JPanel row = new JPanel(new FlowLayout(FlowLayout.CENTER, 2, 0));
        row.setOpaque(false);
        for (JComponent c : items) row.add(c);
        JLabel label = new JLabel(title, SwingConstants.CENTER);
        label.setForeground(new Color(255, 255, 255, 175));
        label.setFont(new Font(Font.SANS_SERIF, Font.PLAIN, 10));
        JPanel middle = new JPanel(new GridBagLayout()); // centres the buttons vertically
        middle.setOpaque(false);
        middle.add(row);
        JPanel g = new JPanel(new BorderLayout(0, 2));
        g.setOpaque(false);
        g.add(middle, BorderLayout.CENTER);
        g.add(label, BorderLayout.SOUTH);
        g.setBorder(new CompoundBorder(new MatteBorder(0, 0, 0, 1, new Color(0x4FAE7B)), new EmptyBorder(0, 6, 0, 6)));
        g.setPreferredSize(new Dimension(g.getPreferredSize().width, 80)); // same height so labels line up
        return g;
    }

    private JComponent homeTab() {
        JPanel p = tabPanel();
        p.add(group("File",
                button("Save", "save", "Save the workbook (Ctrl+S)", e -> saveWorkbook()),
                button("New Sheet", "newsheet", "Add a new worksheet (Shift+F11)", e -> addSheet(nextSheetName())),
                button("Delete Sheet", "deletesheet", "Delete the current worksheet",
                        e -> deleteSheet(tabs.getSelectedIndex()))));

        boldBtn = iconButton("bold", "Bold (Ctrl+B)", e -> toggleStyle(true));
        italicBtn = iconButton("italic", "Italic (Ctrl+I)", e -> toggleStyle(false));
        JLabel sizeLabel = new JLabel("Size");
        sizeLabel.setForeground(Color.WHITE);
        sizeLabel.setFont(new Font(Font.SANS_SERIF, Font.PLAIN, 11));
        sizeCombo.setSelectedItem(13);
        sizeCombo.setFocusable(false);
        sizeCombo.setToolTipText("Font size of the selected cells");
        sizeCombo.addActionListener(e -> {
            if (updatingUi || sizeCombo.getSelectedItem() == null) return;
            final int size = (Integer) sizeCombo.getSelectedItem();
            applyToSelection(st -> st.fontSize = size);
        });
        JPanel sizeBox = new JPanel();
        sizeBox.setOpaque(false);
        sizeBox.setLayout(new BoxLayout(sizeBox, BoxLayout.Y_AXIS));
        sizeBox.add(sizeLabel);
        sizeBox.add(sizeCombo);
        p.add(group("Font", boldBtn, italicBtn, sizeBox));

        alignLeftBtn = iconButton("alignleft", "Align left (Ctrl+L)", e -> align(SwingConstants.LEFT));
        alignCenterBtn = iconButton("aligncenter", "Align center (Ctrl+E)", e -> align(SwingConstants.CENTER));
        alignRightBtn = iconButton("alignright", "Align right (Ctrl+R)", e -> align(SwingConstants.RIGHT));
        p.add(group("Alignment", alignLeftBtn, alignCenterBtn, alignRightBtn));

        fillChip.setPreferredSize(new Dimension(18, 18));
        fillChip.setBackground(currentBackground);
        fillChip.setBorder(BorderFactory.createLineBorder(Color.WHITE, 1));
        fillChip.setToolTipText("Current fill colour");
        JPanel chipBox = new JPanel(new GridBagLayout());
        chipBox.setOpaque(false);
        chipBox.add(fillChip);
        p.add(group("Cells",
                button("Fill Color", "fill", "Pick a background colour for the selected cells", e -> pickFillColor()),
                chipBox,
                button("No Fill", "nofill", "Remove the background colour",
                        e -> applyToSelection(st -> st.background = null))));

        p.add(group("Editing",
                button("AutoSum", "sum", "Sum the numbers above or to the left", e -> autoSum()),
                button("Function", "function", "Browse and insert a function (Shift+F3)", e -> openInsertFunction())));
        return p;
    }

    private JComponent insertTab() {
        JPanel p = tabPanel();
        p.add(group("Illustrations",
                button("Picture", "picture", "Insert a picture from a file", e -> insertPictureFromFile()),
                button("Clipboard", "clipboard", "Paste a picture from the clipboard", e -> insertPictureFromClipboard())));
        p.add(group("Data",
                button("Sample Data", "chart", "Insert date, sample data or a chart placeholder", e -> openInsertDialog()),
                button("Function", "function", "Browse and insert a function (Shift+F3)", e -> openInsertFunction())));
        p.add(group("Code",
                button("VBA Macro", "macro", "Open the macro editor and runner", e -> openMacroIde())));
        protectButton = button("Protect Sheet", "lock", "Protect or unprotect this sheet, optionally with a password",
                e -> toggleProtection());
        p.add(group("Security", protectButton));
        return p;
    }

    private static final Color[] SWATCH_COLORS = {
            new Color(0xE53935), new Color(0xFB8C00), new Color(0xFDD835), new Color(0x43A047),
            new Color(0x1E88E5), new Color(0x8E24AA), new Color(0x212121), new Color(0xFFFFFF)};

    private JComponent drawTab() {
        JPanel p = tabPanel();
        Object[][] defs = {
                {DrawingLayer.Tool.NONE, "cursor", "Select", "Select cells (turns drawing off)"},
                {DrawingLayer.Tool.PEN, "pen", "Pen", "Draw freehand lines"},
                {DrawingLayer.Tool.HIGHLIGHTER, "highlighter", "Highlighter", "Paint translucent highlights over cells"},
                {DrawingLayer.Tool.LINE, "line", "Line", "Straight line - hold Shift for 45° angles"},
                {DrawingLayer.Tool.RECT, "rect", "Rectangle", "Rectangle - hold Shift for a square"},
                {DrawingLayer.Tool.ELLIPSE, "ellipse", "Ellipse", "Ellipse - hold Shift for a circle"},
                {DrawingLayer.Tool.ERASER, "eraser", "Eraser", "Click or drag over a stroke or shape to erase it"}};
        List<JComponent> tools = new ArrayList<>();
        for (Object[] d : defs) {
            DrawingLayer.Tool tool = (DrawingLayer.Tool) d[0];
            RibbonButton b = button((String) d[2], (String) d[1], (String) d[3], e -> setDrawTool(tool));
            toolButtons.put(tool, b);
            tools.add(b);
        }
        toolButtons.get(DrawingLayer.Tool.NONE).setActive(true);
        p.add(group("Tools", tools.toArray(new JComponent[0])));

        JPanel swatchGrid = new JPanel(new java.awt.GridLayout(2, 4, 4, 4));
        swatchGrid.setOpaque(false);
        for (Color c : SWATCH_COLORS) {
            SwatchButton sb = new SwatchButton(c);
            swatches.add(sb);
            swatchGrid.add(sb);
        }
        drawChip.setPreferredSize(new Dimension(22, 22));
        drawChip.setBackground(draw.color);
        drawChip.setBorder(BorderFactory.createLineBorder(Color.WHITE, 2));
        drawChip.setToolTipText("Current drawing colour");
        JPanel chipBox = new JPanel(new GridBagLayout());
        chipBox.setOpaque(false);
        chipBox.add(drawChip);
        p.add(group("Colour", swatchGrid, chipBox,
                button("More", "palette", "Pick any drawing colour", e -> pickDrawColor())));

        JSlider widthSlider = new JSlider(1, 24, 3);
        widthSlider.setOpaque(false);
        widthSlider.setFocusable(false);
        widthSlider.setPreferredSize(new Dimension(120, 22));
        widthSlider.setToolTipText("Brush / line thickness");
        widthLabel.setForeground(Color.WHITE);
        widthLabel.setFont(new Font(Font.SANS_SERIF, Font.PLAIN, 11));
        widthSlider.addChangeListener(e -> {
            draw.width = widthSlider.getValue();
            widthLabel.setText(widthSlider.getValue() + " px");
        });
        JCheckBox fillShapes = new JCheckBox("Fill shapes");
        fillShapes.setOpaque(false);
        fillShapes.setForeground(Color.WHITE);
        fillShapes.setFocusable(false);
        fillShapes.setFont(new Font(Font.SANS_SERIF, Font.PLAIN, 11));
        fillShapes.addActionListener(e -> draw.filled = fillShapes.isSelected());
        JPanel brush = new JPanel();
        brush.setOpaque(false);
        brush.setLayout(new BoxLayout(brush, BoxLayout.Y_AXIS));
        JPanel top = new JPanel(new FlowLayout(FlowLayout.LEFT, 4, 0));
        top.setOpaque(false);
        top.add(new JLabel(Icons.get("pen", 14, Color.WHITE)));
        top.add(widthLabel);
        brush.add(top);
        brush.add(widthSlider);
        brush.add(fillShapes);
        p.add(group("Brush", brush));

        p.add(group("Actions",
                button("Undo", "undo", "Undo the last drawing change (Ctrl+Z)", e -> drawUndo()),
                button("Redo", "redo", "Redo (Ctrl+Y)", e -> drawRedo()),
                button("Clear All", "trash", "Remove every drawing on this sheet (can be undone)", e -> drawClear())));
        return p;
    }

    private JComponent viewTab() {
        JPanel p = tabPanel();
        p.add(group("Zoom",
                button("Zoom In", "zoomin", "Zoom in (Ctrl +, Ctrl+wheel)", e -> zoomBy(10)),
                button("Zoom Out", "zoomout", "Zoom out (Ctrl -, Ctrl+wheel)", e -> zoomBy(-10)),
                button("100%", "zoomreset", "Back to 100% (Ctrl+0)", e -> setZoom(100))));
        p.add(group("Layout",
                button("Reset Columns", "columns", "Restore the default column widths", e -> resetColumnWidths())));
        p.add(group("Help",
                button("Formulas", "help", "Formula reference", e -> showFormulaHelp())));
        return p;
    }

    // ------------------------------------------------------------------
    // Drawing tools
    // ------------------------------------------------------------------

    private void setDrawTool(DrawingLayer.Tool tool) {
        SheetView sv = currentSheet();
        if (tool != DrawingLayer.Tool.NONE && sv != null && sv.model.isProtected()) {
            status("This sheet is protected - unprotect it to draw.");
            Toolkit.getDefaultToolkit().beep();
            tool = DrawingLayer.Tool.NONE;
        }
        draw.tool = tool;
        for (Map.Entry<DrawingLayer.Tool, RibbonButton> en : toolButtons.entrySet()) {
            en.getValue().setActive(en.getKey() == tool);
        }
        for (int i = 0; i < tabs.getTabCount(); i++) {
            Component c = tabs.getComponentAt(i);
            if (c instanceof SheetView) ((SheetView) c).drawing.updateCursor();
        }
        switch (tool) {
            case PEN:
                status("Pen: drag on the sheet to draw. Ctrl+Z undoes, Select (arrow) goes back to cells.");
                break;
            case HIGHLIGHTER:
                status("Highlighter: drag over cells to highlight them.");
                break;
            case LINE:
                status("Line: drag from start to end (hold Shift for 45° angles).");
                break;
            case RECT:
                status("Rectangle: drag a corner to corner (Shift = square).");
                break;
            case ELLIPSE:
                status("Ellipse: drag across its bounds (Shift = circle).");
                break;
            case ERASER:
                status("Eraser: click or drag across a stroke or shape to remove it.");
                break;
            default:
                break;
        }
    }

    private void selectDrawColor(Color c) {
        draw.color = c;
        drawChip.setBackground(c);
        for (SwatchButton sb : swatches) sb.repaint();
    }

    private void pickDrawColor() {
        Color c = javax.swing.JColorChooser.showDialog(this, "Drawing colour", draw.color);
        if (c != null) selectDrawColor(c);
    }

    private void drawUndo() {
        SheetView sv = currentSheet();
        if (sv == null) return;
        status(sv.drawing.undo() ? "Undid the last drawing change" : "Nothing to undo");
    }

    private void drawRedo() {
        SheetView sv = currentSheet();
        if (sv == null) return;
        status(sv.drawing.redo() ? "Redid the drawing change" : "Nothing to redo");
    }

    private void drawClear() {
        SheetView sv = currentSheet();
        if (sv == null || !guardEditable(sv)) return;
        status(sv.drawing.clearAll() ? "All drawings removed (Undo brings them back)" : "There are no drawings on this sheet");
    }

    /** Round colour button for the quick palette. */
    private final class SwatchButton extends JComponent {
        private final Color color;
        private boolean hover;

        SwatchButton(Color color) {
            this.color = color;
            setPreferredSize(new Dimension(20, 20));
            setCursor(Cursor.getPredefinedCursor(Cursor.HAND_CURSOR));
            setToolTipText("Drawing colour");
            addMouseListener(new MouseAdapter() {
                @Override
                public void mouseClicked(MouseEvent e) {
                    selectDrawColor(SwatchButton.this.color);
                }

                @Override
                public void mouseEntered(MouseEvent e) {
                    hover = true;
                    repaint();
                }

                @Override
                public void mouseExited(MouseEvent e) {
                    hover = false;
                    repaint();
                }
            });
        }

        @Override
        protected void paintComponent(Graphics g) {
            Graphics2D g2 = (Graphics2D) g.create();
            g2.setRenderingHint(RenderingHints.KEY_ANTIALIASING, RenderingHints.VALUE_ANTIALIAS_ON);
            int inset = hover ? 1 : 3;
            g2.setColor(color);
            g2.fillOval(inset, inset, getWidth() - 2 * inset, getHeight() - 2 * inset);
            g2.setColor(color.equals(draw.color) ? Color.WHITE : new Color(255, 255, 255, 140));
            g2.setStroke(new BasicStroke(color.equals(draw.color) ? 2.4f : 1f));
            g2.drawOval(inset, inset, getWidth() - 2 * inset - 1, getHeight() - 2 * inset - 1);
            g2.dispose();
        }
    }

    // ------------------------------------------------------------------
    // Ribbon widgets
    // ------------------------------------------------------------------

    private RibbonButton button(String text, String icon, String tip, java.awt.event.ActionListener action) {
        RibbonButton b = new RibbonButton(text, icon, tip, 24);
        b.addActionListener(action);
        return b;
    }

    private RibbonButton iconButton(String icon, String tip, java.awt.event.ActionListener action) {
        RibbonButton b = new RibbonButton("", icon, tip, 20);
        b.addActionListener(action);
        return b;
    }

    /** Ribbon tab (Home / Insert / Draw / View). The selected one merges with the ribbon body. */
    private final class RibbonTab extends JButton {
        private final String name;

        RibbonTab(String name) {
            super(name);
            this.name = name;
            setUI(new BasicButtonUI());
            setRolloverEnabled(true);
            setFocusable(false);
            setFocusPainted(false);
            setBorderPainted(false);
            setContentAreaFilled(false);
            setOpaque(false);
            setForeground(Color.WHITE);
            setFont(new Font(Font.SANS_SERIF, Font.BOLD, 12));
            setBorder(new EmptyBorder(6, 16, 6, 16));
            setCursor(Cursor.getPredefinedCursor(Cursor.HAND_CURSOR));
            addActionListener(e -> selectRibbonTab(name));
        }

        @Override
        protected void paintComponent(Graphics g) {
            Graphics2D g2 = (Graphics2D) g.create();
            g2.setRenderingHint(RenderingHints.KEY_ANTIALIASING, RenderingHints.VALUE_ANTIALIAS_ON);
            if (name.equals(ribbonSelected)) {
                g2.setColor(GREEN);
                g2.fillRoundRect(0, 0, getWidth(), getHeight() + 10, 10, 10);
            } else if (getModel().isRollover()) {
                g2.setColor(new Color(255, 255, 255, 40));
                g2.fillRoundRect(0, 2, getWidth(), getHeight(), 8, 8);
            }
            g2.dispose();
            super.paintComponent(g);
        }
    }

    /** Flat ribbon button: icon above the label, hover / pressed highlight and an "on" state. */
    private static final class RibbonButton extends JButton {
        private final int iconSize;
        private boolean active;

        RibbonButton(String text, String icon, String tip, int iconSize) {
            super(text);
            this.iconSize = iconSize;
            setUI(new BasicButtonUI());
            setRolloverEnabled(true);
            setToolTipText(tip);
            setForeground(Color.WHITE);
            setFont(new Font(Font.SANS_SERIF, Font.PLAIN, 11));
            setFocusPainted(false);
            setFocusable(false); // keeps keyboard focus in the grid
            setBorderPainted(false);
            setContentAreaFilled(false);
            setOpaque(false);
            boolean iconOnly = text == null || text.isEmpty();
            setBorder(iconOnly ? new EmptyBorder(7, 7, 7, 7) : new EmptyBorder(5, 8, 3, 8));
            if (icon != null) setIcon(Icons.get(icon, iconSize, Color.WHITE));
            setVerticalTextPosition(SwingConstants.BOTTOM);
            setHorizontalTextPosition(SwingConstants.CENTER);
            setIconTextGap(3);
            setCursor(Cursor.getPredefinedCursor(Cursor.HAND_CURSOR));
        }

        void setActive(boolean on) {
            if (active != on) {
                active = on;
                repaint();
            }
        }

        void setIconName(String icon) {
            setIcon(Icons.get(icon, iconSize, Color.WHITE));
        }

        @Override
        protected void paintComponent(Graphics g) {
            ButtonModel m = getModel();
            Graphics2D g2 = (Graphics2D) g.create();
            g2.setRenderingHint(RenderingHints.KEY_ANTIALIASING, RenderingHints.VALUE_ANTIALIAS_ON);
            if (active) {
                g2.setColor(new Color(255, 255, 255, 70));
                g2.fillRoundRect(0, 0, getWidth(), getHeight(), 10, 10);
                g2.setColor(new Color(255, 255, 255, 150));
                g2.drawRoundRect(0, 0, getWidth() - 1, getHeight() - 1, 10, 10);
            }
            if (m.isPressed()) {
                g2.setColor(new Color(0, 0, 0, 60));
                g2.fillRoundRect(0, 0, getWidth(), getHeight(), 10, 10);
            } else if (m.isRollover()) {
                g2.setColor(new Color(255, 255, 255, 45));
                g2.fillRoundRect(0, 0, getWidth(), getHeight(), 10, 10);
            }
            g2.dispose();
            super.paintComponent(g);
        }
    }

    // ==================================================================
    // Formula bar and status bar
    // ==================================================================

    private JComponent buildFormulaBar() {
        JPanel bar = new JPanel(new BorderLayout(6, 0));
        bar.setBackground(Color.WHITE);
        bar.setBorder(new CompoundBorder(new MatteBorder(0, 0, 1, 0, GRID), new EmptyBorder(4, 6, 4, 6)));

        cellBadge.setPreferredSize(new Dimension(90, 26));
        cellBadge.setOpaque(true);
        cellBadge.setBackground(LIGHT_GRAY);
        cellBadge.setFont(new Font(Font.SANS_SERIF, Font.BOLD, 13));
        cellBadge.setBorder(BorderFactory.createLineBorder(GRID));

        JLabel fx = new JLabel("fx");
        fx.setFont(new Font(Font.SERIF, Font.BOLD | Font.ITALIC, 16));
        fx.setForeground(GREEN);
        fx.setBorder(new EmptyBorder(0, 6, 0, 6));
        fx.setToolTipText("Insert Function...");
        fx.setCursor(Cursor.getPredefinedCursor(Cursor.HAND_CURSOR));
        fx.addMouseListener(new MouseAdapter() {
            @Override
            public void mouseClicked(MouseEvent e) {
                openInsertFunction();
            }
        });

        JPanel left = new JPanel(new BorderLayout(6, 0));
        left.setOpaque(false);
        left.add(cellBadge, BorderLayout.WEST);
        left.add(fx, BorderLayout.CENTER);

        formulaField.setFont(new Font(Font.SANS_SERIF, Font.PLAIN, 13));
        formulaField.setBorder(new CompoundBorder(BorderFactory.createLineBorder(GRID), new EmptyBorder(2, 6, 2, 6)));
        formulaField.addActionListener(e -> commitFormulaBar());
        formulaField.getInputMap().put(KeyStroke.getKeyStroke(KeyEvent.VK_ESCAPE, 0), "revert");
        formulaField.getActionMap().put("revert", new AbstractAction() {
            @Override
            public void actionPerformed(ActionEvent e) {
                syncFormulaBar();
                SheetView sv = currentSheet();
                if (sv != null) sv.table.requestFocusInWindow();
            }
        });

        bar.add(left, BorderLayout.WEST);
        bar.add(formulaField, BorderLayout.CENTER);
        return bar;
    }

    private JComponent buildStatusBar() {
        JPanel bar = new JPanel(new BorderLayout());
        bar.setBackground(LIGHT_GRAY);
        bar.setBorder(new CompoundBorder(new MatteBorder(1, 0, 0, 0, GRID), new EmptyBorder(3, 10, 3, 10)));
        bar.add(statusLabel, BorderLayout.WEST);
        protLabel.setForeground(new Color(0xC00000));
        protLabel.setFont(protLabel.getFont().deriveFont(Font.BOLD));
        bar.add(protLabel, BorderLayout.CENTER);
        JPanel east = new JPanel(new FlowLayout(FlowLayout.RIGHT, 6, 0));
        east.setOpaque(false);
        east.add(statsLabel);
        JButton minus = zoomButton("-", "Zoom out");
        minus.addActionListener(e -> zoomBy(-10));
        JButton plus = zoomButton("+", "Zoom in");
        plus.addActionListener(e -> zoomBy(10));
        zoomSlider.setOpaque(false);
        zoomSlider.setFocusable(false);
        zoomSlider.setPreferredSize(new Dimension(120, 20));
        zoomSlider.addChangeListener(e -> {
            if (!updatingUi) setZoom(zoomSlider.getValue());
        });
        zoomLabel.setPreferredSize(new Dimension(44, 18));
        zoomLabel.setHorizontalAlignment(SwingConstants.RIGHT);
        zoomLabel.setToolTipText("Click to reset to 100%");
        zoomLabel.addMouseListener(new MouseAdapter() {
            @Override
            public void mouseClicked(MouseEvent e) {
                setZoom(100);
            }
        });
        east.add(minus);
        east.add(zoomSlider);
        east.add(plus);
        east.add(zoomLabel);
        bar.add(east, BorderLayout.EAST);
        return bar;
    }

    private JButton zoomButton(String text, String tip) {
        JButton b = new JButton(text);
        b.setMargin(new java.awt.Insets(0, 4, 0, 4));
        b.setPreferredSize(new Dimension(22, 20));
        b.setFocusable(false);
        b.setToolTipText(tip);
        return b;
    }

    // ------------------------------------------------------------------
    // Zoom and Insert Function
    // ------------------------------------------------------------------

    private void setZoom(int percent) {
        SheetView sv = currentSheet();
        if (sv == null) return;
        sv.setZoom(percent);
        updateZoomUi(sv);
        markDirty();
        status("Zoom " + sv.model.getZoomPercent() + "%");
    }

    private void zoomBy(int delta) {
        SheetView sv = currentSheet();
        if (sv != null) setZoom(sv.model.getZoomPercent() + delta);
    }

    private void updateZoomUi(SheetView sv) {
        int z = sv.model.getZoomPercent();
        updatingUi = true;
        zoomSlider.setValue(z);
        updatingUi = false;
        zoomLabel.setText(z + "%");
    }

    private void openInsertFunction() {
        SheetView sv = currentSheet();
        if (sv == null || !guardEditable(sv)) return;
        new InsertFunctionDialog(this, this::insertFunction).setVisible(true);
    }

    /** Puts NAME() into the formula bar (after an existing '=' formula, at the caret) with the caret inside. */
    private void insertFunction(String name) {
        String text = formulaField.getText();
        int caret = formulaField.getCaretPosition();
        String call = name + "()";
        int newCaret;
        if (!text.startsWith("=")) {
            text = "=" + call;
            newCaret = name.length() + 2;
        } else {
            caret = Math.max(1, Math.min(caret, text.length()));
            text = text.substring(0, caret) + call + text.substring(caret);
            newCaret = caret + name.length() + 1;
        }
        formulaField.setText(text);
        formulaField.requestFocusInWindow();
        formulaField.setCaretPosition(newCaret);
        status("Inserted " + name + " - fill in the arguments and press Enter");
    }

    private void status(String message) {
        statusLabel.setText(message);
    }

    /** Pushes the current selection into the badge, formula field, size box and stats label. */
    private void syncFormulaBar() {
        SheetView sv = currentSheet();
        if (sv == null) return;
        updateProtectUi(sv);
        updateZoomUi(sv);
        JTable t = sv.table;
        int r = t.getSelectionModel().getLeadSelectionIndex();
        int c = t.getColumnModel().getSelectionModel().getLeadSelectionIndex();
        if (!sv.model.inBounds(r, c)) {
            cellBadge.setText("");
            formulaField.setText("");
            statsLabel.setText("");
            return;
        }
        int[] rows = t.getSelectedRows();
        int[] cols = t.getSelectedColumns();
        if (rows.length * cols.length > 1) {
            cellBadge.setText(FormulaEngine.cellRef(rows[0], cols[0]) + ":"
                    + FormulaEngine.cellRef(rows[rows.length - 1], cols[cols.length - 1]));
        } else {
            cellBadge.setText(FormulaEngine.cellRef(r, c));
        }
        formulaField.setText(sv.model.getInput(r, c));
        formulaField.setCaretPosition(0);

        CellStyle st = sv.model.getStyle(r, c);
        updatingUi = true;
        sizeCombo.setSelectedItem(st != null && st.fontSize > 0 ? st.fontSize : 13);
        updatingUi = false;
        if (boldBtn != null) {
            boldBtn.setActive(st != null && st.bold);
            italicBtn.setActive(st != null && st.italic);
            alignLeftBtn.setActive(st != null && st.alignment == SwingConstants.LEFT);
            alignCenterBtn.setActive(st != null && st.alignment == SwingConstants.CENTER);
            alignRightBtn.setActive(st != null && st.alignment == SwingConstants.RIGHT);
        }

        // Excel-like statistics for multi-cell selections
        int count = 0;
        double sum = 0;
        if (rows.length * cols.length > 1) {
            for (int rr : rows) {
                for (int cc : cols) {
                    Object v = sv.model.getComputedValue(rr, cc);
                    if (v instanceof Double) {
                        count++;
                        sum += (Double) v;
                    }
                }
            }
        }
        statsLabel.setText(count > 0
                ? "Count: " + count + "     Sum: " + FormulaEngine.formatNumber(sum)
                + "     Average: " + FormulaEngine.formatNumber(sum / count)
                : "");
    }

    /** Enter in the formula bar: write the text into the active cell and move down. */
    private void commitFormulaBar() {
        SheetView sv = currentSheet();
        if (sv == null) return;
        int r = sv.table.getSelectionModel().getLeadSelectionIndex();
        int c = sv.table.getColumnModel().getSelectionModel().getLeadSelectionIndex();
        if (!sv.model.inBounds(r, c)) return;
        if (!guardEditable(sv)) return;
        if (sv.table.isEditing()) sv.table.getCellEditor().cancelCellEditing();
        sv.model.setInput(r, c, formulaField.getText());
        if (r + 1 < sv.model.getRowCount()) sv.table.changeSelection(r + 1, c, false, false);
        sv.table.requestFocusInWindow();
    }

    // ==================================================================
    // Menu bar
    // ==================================================================

    private JMenuItem menuItem(String text, int key, int modifiers, Runnable action) {
        JMenuItem item = new JMenuItem(text);
        String ic = menuIconFor(text);
        if (ic != null) item.setIcon(Icons.get(ic, 16, new Color(0x444444)));
        if (key != 0) item.setAccelerator(KeyStroke.getKeyStroke(key, modifiers));
        item.addActionListener(e -> action.run());
        return item;
    }

    private static String menuIconFor(String text) {
        switch (text) {
            case "New Workbook":
            case "New Sheet":
                return "newsheet";
            case "Open...":
                return "open";
            case "Save":
            case "Save As...":
                return "save";
            case "Set Open Password...":
            case "Protect Sheet...":
                return "lock";
            case "Unprotect Sheet...":
                return "unlock";
            case "Clear Contents":
            case "Delete Selected Picture":
                return "trash";
            case "Zoom In":
                return "zoomin";
            case "Zoom Out":
                return "zoomout";
            case "Zoom 100%":
                return "zoomreset";
            case "Reset Column Widths":
                return "columns";
            case "Function...":
                return "function";
            case "Picture from File...":
                return "picture";
            case "Picture from Clipboard":
                return "clipboard";
            case "Date / Sample Data / Chart...":
                return "chart";
            case "Formula Reference":
                return "help";
            default:
                return null;
        }
    }

    private JMenuBar buildMenuBar() {
        int ctrl = InputEvent.CTRL_DOWN_MASK;
        JMenuBar bar = new JMenuBar();

        JMenu file = new JMenu("File");
        file.add(menuItem("New Workbook", KeyEvent.VK_N, ctrl, this::newWorkbook));
        file.add(menuItem("Open...", KeyEvent.VK_O, ctrl, this::openWorkbook));
        file.add(menuItem("Save", KeyEvent.VK_S, ctrl, () -> saveWorkbook()));
        file.add(menuItem("Save As...", KeyEvent.VK_S, ctrl | InputEvent.SHIFT_DOWN_MASK, () -> saveWorkbookAs()));
        file.addSeparator();
        file.add(menuItem("Set Open Password...", 0, 0, this::setOpenPassword));
        file.addSeparator();
        file.add(menuItem("Import CSV...", 0, 0, this::importCsv));
        file.add(menuItem("Export CSV...", 0, 0, this::exportCsv));
        file.addSeparator();
        file.add(menuItem("Exit", 0, 0, this::exitApp));

        JMenu edit = new JMenu("Edit");
        edit.add(menuItem("Clear Contents", 0, 0, this::clearSelection));
        edit.add(menuItem("Clear Formatting", 0, 0, () -> applyToSelection(CellStyle::reset)));
        edit.addSeparator();
        edit.add(menuItem("Delete Selected Picture", 0, 0, this::deleteSelectedPicture));

        JMenu view = new JMenu("View");
        view.add(menuItem("Zoom In", KeyEvent.VK_EQUALS, ctrl, () -> zoomBy(10)));
        view.add(menuItem("Zoom Out", KeyEvent.VK_MINUS, ctrl, () -> zoomBy(-10)));
        view.add(menuItem("Zoom 100%", KeyEvent.VK_0, ctrl, () -> setZoom(100)));
        view.addSeparator();
        for (int z : new int[]{50, 75, 100, 125, 150, 200}) {
            view.add(menuItem(z + "%", 0, 0, () -> setZoom(z)));
        }
        view.addSeparator();
        view.add(menuItem("Reset Column Widths", 0, 0, this::resetColumnWidths));

        JMenu insert = new JMenu("Insert");
        insert.add(menuItem("New Sheet", KeyEvent.VK_F11, InputEvent.SHIFT_DOWN_MASK, () -> addSheet(nextSheetName())));
        insert.add(menuItem("Function...", KeyEvent.VK_F3, InputEvent.SHIFT_DOWN_MASK, this::openInsertFunction));
        insert.addSeparator();
        insert.add(menuItem("Picture from File...", 0, 0, this::insertPictureFromFile));
        insert.add(menuItem("Picture from Clipboard", 0, 0, this::insertPictureFromClipboard));
        insert.addSeparator();
        insert.add(menuItem("Date / Sample Data / Chart...", 0, 0, this::openInsertDialog));

        JMenu review = new JMenu("Review");
        review.add(menuItem("Protect Sheet...", 0, 0, this::protectSheetDialog));
        review.add(menuItem("Unprotect Sheet...", 0, 0, this::unprotectSheetDialog));

        JMenu help = new JMenu("Help");
        help.add(menuItem("Formula Reference", 0, 0, this::showFormulaHelp));

        bar.add(file);
        bar.add(edit);
        bar.add(view);
        bar.add(insert);
        bar.add(review);
        bar.add(help);
        return bar;
    }

    private JMenuItem ctxItem(String text, String icon, Runnable action) {
        JMenuItem item = new JMenuItem(text, Icons.get(icon, 16, new Color(0x444444)));
        item.addActionListener(e -> action.run());
        return item;
    }

    /** Right-click menu for the grid. */
    private void showCellMenu(SheetView sv, MouseEvent e) {
        JTable t = sv.table;
        int r = t.rowAtPoint(e.getPoint());
        int c = t.columnAtPoint(e.getPoint());
        if (r < 0 || c < 0) return;
        if (!t.isCellSelected(r, c)) t.changeSelection(r, c, false, false);
        t.requestFocusInWindow();
        JPopupMenu m = new JPopupMenu();
        m.add(ctxItem("Insert Function...", "function", this::openInsertFunction));
        m.add(ctxItem("Clear Contents", "trash", this::clearSelection));
        m.addSeparator();
        m.add(ctxItem("Bold", "bold", () -> toggleStyle(true)));
        m.add(ctxItem("Italic", "italic", () -> toggleStyle(false)));
        m.add(ctxItem("Align Left", "alignleft", () -> align(SwingConstants.LEFT)));
        m.add(ctxItem("Align Center", "aligncenter", () -> align(SwingConstants.CENTER)));
        m.add(ctxItem("Align Right", "alignright", () -> align(SwingConstants.RIGHT)));
        m.addSeparator();
        m.add(ctxItem("Fill Color...", "fill", this::pickFillColor));
        m.add(ctxItem("No Fill", "nofill", () -> applyToSelection(st -> st.background = null)));
        m.addSeparator();
        final int col = c;
        m.add(ctxItem("Fit Column " + FormulaEngine.columnLabel(c) + " to Content", "columns", () -> sv.autoFitColumn(col)));
        m.add(ctxItem("Insert Picture...", "picture", this::insertPictureFromFile));
        m.show(t, e.getX(), e.getY());
    }

    private void resetColumnWidths() {
        SheetView sv = currentSheet();
        if (sv == null) return;
        sv.model.resetColumnWidths();
        sv.applyZoom();
        markDirty();
        status("Column widths reset");
    }

    private void showFormulaHelp() {
        StringBuilder sb = new StringBuilder();
        sb.append("Formulas start with '='.\n\n")
          .append("Operators:   +  -  *  /  ^  %  &  = <> < > <= >=  and parentheses\n")
          .append("References:  =A1+B2*5      Ranges: =SUM(A1:A10)\n")
          .append("Text goes in quotes: =\"Hello \"&A1      TRUE and FALSE are values\n\n")
          .append("Click the fx button (or Insert > Function) to browse functions.\n\n");
        String cat = "";
        for (FunctionCatalog.Entry e : FunctionCatalog.all()) {
            if (!e.category.equals(cat)) {
                cat = e.category;
                sb.append("\n== ").append(cat).append(" ==\n");
            }
            sb.append("  ").append(e.syntax).append("\n");
        }
        sb.append("\nErrors:\n")
          .append("  #DIV/0!   division by zero\n")
          .append("  #SYNTAX!  malformed formula or wrong number of arguments\n")
          .append("  #NAME?    unknown function or name\n")
          .append("  #N/A      lookup value not found\n")
          .append("  #ERROR!   bad reference, text used as a number, circular reference\n\n")
          .append("View: Ctrl+mouse wheel or Ctrl +/- / 0 zoom; drag a column header edge to resize,\n")
          .append("double-click the edge to fit the content.\n")
          .append("Shortcuts: Ctrl+B bold, Ctrl+I italic, Ctrl+L/E/R align, Delete clear, F2 edit cell");
        JTextArea text = new JTextArea(sb.toString(), 24, 62);
        text.setEditable(false);
        text.setCaretPosition(0);
        text.setFont(new Font(Font.MONOSPACED, Font.PLAIN, 12));
        JOptionPane.showMessageDialog(this, new JScrollPane(text), "Formula Reference", JOptionPane.INFORMATION_MESSAGE);
    }

    // ==================================================================
    // Sheets (tabs)
    // ==================================================================

    private String nextSheetName() {
        return "Sheet" + (++sheetCounter);
    }

    private SheetView currentSheet() {
        Component c = tabs.getSelectedComponent();
        return c instanceof SheetView ? (SheetView) c : null;
    }

    private SheetView addSheet(String name) {
        return addSheet(name, new SpreadsheetModel());
    }

    private SheetView addSheet(String name, SpreadsheetModel model) {
        SheetView sv = new SheetView(model);
        tabs.addTab(name, sv);
        int idx = tabs.getTabCount() - 1;
        TabHeader header = new TabHeader(name);
        header.setLocked(model.isProtected());
        tabs.setTabComponentAt(idx, header);
        tabs.setSelectedIndex(idx);
        SwingUtilities.invokeLater(() -> sv.table.requestFocusInWindow());
        status("Added " + name);
        markDirty();
        return sv;
    }

    private void deleteSheet(int index) {
        if (index < 0) return;
        if (tabs.getTabCount() <= 1) {
            JOptionPane.showMessageDialog(this, "A workbook must contain at least one sheet.",
                    "Delete Sheet", JOptionPane.WARNING_MESSAGE);
            return;
        }
        Component comp = tabs.getComponentAt(index);
        if (comp instanceof SheetView && ((SheetView) comp).model.isProtected()) {
            JOptionPane.showMessageDialog(this, "This sheet is protected. Unprotect it before deleting it.",
                    "Delete Sheet", JOptionPane.WARNING_MESSAGE);
            return;
        }
        String title = tabs.getTitleAt(index);
        int answer = JOptionPane.showConfirmDialog(this,
                "Delete \"" + title + "\"? This cannot be undone.", "Delete Sheet",
                JOptionPane.OK_CANCEL_OPTION, JOptionPane.WARNING_MESSAGE);
        if (answer == JOptionPane.OK_OPTION) {
            tabs.removeTabAt(index);
            status("Deleted " + title);
            markDirty();
        }
    }

    private void renameSheet(int index) {
        if (index < 0) return;
        Object input = JOptionPane.showInputDialog(this, "Sheet name:", "Rename Sheet",
                JOptionPane.PLAIN_MESSAGE, null, null, tabs.getTitleAt(index));
        if (input == null) return;
        String name = input.toString().trim();
        if (name.isEmpty()) return;
        tabs.setTitleAt(index, name);
        Component header = tabs.getTabComponentAt(index);
        if (header instanceof TabHeader) ((TabHeader) header).setTitle(name);
        markDirty();
    }

    /** Custom tab title with a close button, a rename-on-double-click and a context menu. */
    private final class TabHeader extends JPanel {
        private final JLabel title;
        private final JLabel lockLabel = new JLabel(new LockIcon());

        TabHeader(String name) {
            super(new FlowLayout(FlowLayout.LEFT, 4, 0));
            setOpaque(false);
            lockLabel.setVisible(false);
            lockLabel.setToolTipText("This sheet is protected");
            add(lockLabel);
            title = new JLabel(name);
            add(title);

            JButton close = new JButton("\u00D7");
            close.setFocusable(false);
            close.setBorder(new EmptyBorder(0, 3, 0, 3));
            close.setContentAreaFilled(false);
            close.setFont(close.getFont().deriveFont(Font.BOLD, 14f));
            close.setForeground(Color.GRAY);
            close.setToolTipText("Close sheet");
            close.addActionListener(e -> deleteSheet(tabs.indexOfTabComponent(TabHeader.this)));
            close.addMouseListener(new MouseAdapter() {
                @Override
                public void mouseEntered(MouseEvent e) {
                    close.setForeground(new Color(0xC00000));
                }

                @Override
                public void mouseExited(MouseEvent e) {
                    close.setForeground(Color.GRAY);
                }
            });
            add(close);

            MouseAdapter ma = new MouseAdapter() {
                @Override
                public void mousePressed(MouseEvent e) {
                    int i = tabs.indexOfTabComponent(TabHeader.this);
                    if (i >= 0) tabs.setSelectedIndex(i);
                    if (e.isPopupTrigger()) showMenu(e);
                }

                @Override
                public void mouseReleased(MouseEvent e) {
                    if (e.isPopupTrigger()) showMenu(e);
                }

                @Override
                public void mouseClicked(MouseEvent e) {
                    if (e.getClickCount() == 2 && SwingUtilities.isLeftMouseButton(e)) {
                        renameSheet(tabs.indexOfTabComponent(TabHeader.this));
                    }
                }
            };
            addMouseListener(ma);
            title.addMouseListener(ma);
        }

        void setTitle(String name) {
            title.setText(name);
        }

        void setLocked(boolean locked) {
            lockLabel.setVisible(locked);
            revalidate();
            repaint();
        }

        private void showMenu(MouseEvent e) {
            JPopupMenu menu = new JPopupMenu();
            JMenuItem rename = new JMenuItem("Rename...");
            rename.addActionListener(a -> renameSheet(tabs.indexOfTabComponent(TabHeader.this)));
            JMenuItem delete = new JMenuItem("Delete");
            delete.addActionListener(a -> deleteSheet(tabs.indexOfTabComponent(TabHeader.this)));
            menu.add(rename);
            menu.add(delete);
            menu.show(e.getComponent(), e.getX(), e.getY());
        }
    }

    // ==================================================================
    // One worksheet view (table + headers)
    // ==================================================================

    private final class SheetView extends JPanel {
        final SpreadsheetModel model;
        final JTable table;
        final JList<String> rowHeader;
        ImageLayer layer; // pictures floating above the grid
        DrawingLayer drawing; // pen / shapes painted above the pictures
        private final SpreadsheetCellRenderer.ColumnHeaderRenderer colHeaderRenderer =
                new SpreadsheetCellRenderer.ColumnHeaderRenderer();
        private SpreadsheetCellRenderer.RowHeaderRenderer rowHeaderRenderer;
        final SpreadsheetCellRenderer cellRenderer;
        final FormulaCellEditor cellEditor;
        private final JScrollPane scroll;
        private boolean applying = false;
        private static final int BASE_FONT = 13, BASE_ROW = 22, BASE_ROW_HEADER = 46, BASE_HEADER = 24;

        SheetView(SpreadsheetModel sheetModel) {
            super(new BorderLayout());
            this.model = sheetModel;

            table = new JTable(model) {
                @Override
                protected void paintComponent(Graphics g) {
                    super.paintComponent(g);
                    if (SheetView.this.layer != null) SheetView.this.layer.paint(g);
                    if (SheetView.this.drawing != null) SheetView.this.drawing.paint(g);
                }

                @Override
                protected void processMouseEvent(MouseEvent e) {
                    if (SheetView.this.drawing != null && SheetView.this.drawing.handleMouse(e)) return;
                    if (SheetView.this.layer != null && SheetView.this.layer.handleMouse(e)) return;
                    super.processMouseEvent(e);
                }

                @Override
                protected void processMouseMotionEvent(MouseEvent e) {
                    if (SheetView.this.drawing != null && SheetView.this.drawing.handleMotion(e)) return;
                    if (SheetView.this.layer != null && SheetView.this.layer.handleMotion(e)) return;
                    super.processMouseMotionEvent(e);
                }

                @Override
                public boolean editCellAt(int row, int column, java.util.EventObject e) {
                    if (SheetView.this.model.isProtected()) {
                        guardEditable(SheetView.this);
                        return false;
                    }
                    return super.editCellAt(row, column, e);
                }

                /** Typing over a selected cell replaces its content, like Excel. */
                @Override
                protected boolean processKeyBinding(KeyStroke ks, KeyEvent e, int condition, boolean pressed) {
                    boolean wasEditing = isEditing();
                    boolean handled = super.processKeyBinding(ks, e, condition, pressed);
                    if (!wasEditing && isEditing() && e.getID() == KeyEvent.KEY_PRESSED
                            && e.getKeyChar() != KeyEvent.CHAR_UNDEFINED
                            && getEditorComponent() instanceof JTextField) {
                        ((JTextField) getEditorComponent()).selectAll();
                    }
                    return handled;
                }
            };
            table.setFont(new Font(Font.SANS_SERIF, Font.PLAIN, 13));
            table.setRowHeight(22);
            table.setAutoResizeMode(JTable.AUTO_RESIZE_OFF);
            table.setCellSelectionEnabled(true);
            table.setSelectionMode(ListSelectionModel.MULTIPLE_INTERVAL_SELECTION);
            table.setShowGrid(false);
            table.setIntercellSpacing(new Dimension(0, 0));
            table.getTableHeader().setReorderingAllowed(false);
            table.getTableHeader().setDefaultRenderer(colHeaderRenderer);
            table.getTableHeader().setPreferredSize(new Dimension(100, 24));
            cellRenderer = new SpreadsheetCellRenderer(model);
            cellEditor = new FormulaCellEditor(model, table.getFont());
            table.setDefaultRenderer(Object.class, cellRenderer);
            table.setDefaultEditor(Object.class, cellEditor);
            layer = new ImageLayer(table, model, SpreadsheetApp.this::markDirty);
            drawing = new DrawingLayer(table, model, draw, SpreadsheetApp.this::markDirty,
                    () -> status("This sheet is protected - unprotect it to draw."));

            String[] labels = new String[model.getRowCount()];
            for (int i = 0; i < labels.length; i++) labels[i] = String.valueOf(i + 1);
            rowHeader = new JList<>(labels);
            rowHeader.setFixedCellHeight(table.getRowHeight());
            rowHeader.setFixedCellWidth(46);
            rowHeaderRenderer = new SpreadsheetCellRenderer.RowHeaderRenderer(table);
            rowHeader.setCellRenderer(rowHeaderRenderer);
            rowHeader.setBackground(SpreadsheetCellRenderer.HEADER_BG);
            rowHeader.addMouseListener(new MouseAdapter() {
                @Override
                public void mousePressed(MouseEvent e) {
                    int i = rowHeader.locationToIndex(e.getPoint());
                    if (i < 0) return;
                    table.setRowSelectionInterval(i, i);
                    table.setColumnSelectionInterval(model.getColumnCount() - 1, 0); // active cell = column A
                    table.requestFocusInWindow();
                }
            });

            scroll = new JScrollPane(table);
            scroll.setRowHeaderView(rowHeader);
            JPanel corner = new JPanel();
            corner.setBackground(SpreadsheetCellRenderer.HEADER_BG);
            scroll.setCorner(JScrollPane.UPPER_LEFT_CORNER, corner);
            scroll.setBorder(BorderFactory.createEmptyBorder());
            scroll.getVerticalScrollBar().setUnitIncrement(22);
            add(scroll, BorderLayout.CENTER);

            installKeyBindings();
            installZoomAndResize();
            installHover();
            applyZoom();
            drawing.updateCursor();

            ListSelectionListener l = e -> {
                if (e.getValueIsAdjusting()) return;
                if (SheetView.this == currentSheet()) syncFormulaBar();
                rowHeader.repaint();
                table.getTableHeader().repaint();
            };
            table.getSelectionModel().addListSelectionListener(l);
            table.getColumnModel().getSelectionModel().addListSelectionListener(l);

            // Reactive formula bar: whenever the model changes (typing, macros, insert) refresh it
            model.addTableModelListener(e -> {
                markDirty();
                if (SheetView.this == currentSheet()) syncFormulaBar();
            });

            table.changeSelection(0, 0, false, false);
        }

        double factor() {
            return model.getZoomPercent() / 100.0;
        }

        private int scaled(int base) {
            return Math.max(1, (int) Math.round(base * factor()));
        }

        void setZoom(int percent) {
            Rectangle vis = scroll.getViewport().getViewRect();
            double old = factor();
            model.setZoomPercent(percent);
            double ratio = factor() / old;
            applyZoom();
            if (ratio != 1.0) {
                scroll.getViewport().setViewPosition(new Point(
                        (int) Math.round(vis.x * ratio), (int) Math.round(vis.y * ratio)));
            }
        }

        /** Re-lays out fonts, row height and column widths for the model's zoom level and stored widths. */
        void applyZoom() {
            applying = true;
            try {
                double f = factor();
                Font font = new Font(Font.SANS_SERIF, Font.PLAIN, 13).deriveFont((float) (BASE_FONT * f));
                table.setFont(font);
                cellRenderer.setZoom(f);
                cellEditor.setFont(font);
                table.setRowHeight(scaled(BASE_ROW));
                for (int c = 0; c < model.getColumnCount(); c++) {
                    TableColumn col = table.getColumnModel().getColumn(c);
                    int w = scaled(model.getColumnWidth(c));
                    col.setPreferredWidth(w);
                    col.setWidth(w);
                }
                rowHeader.setFixedCellHeight(table.getRowHeight());
                rowHeader.setFixedCellWidth(scaled(BASE_ROW_HEADER));
                table.getTableHeader().setPreferredSize(new Dimension(100, scaled(BASE_HEADER)));
                table.getTableHeader().resizeAndRepaint();
                table.revalidate();
                scroll.getVerticalScrollBar().setUnitIncrement(table.getRowHeight());
                scroll.revalidate();
                scroll.repaint();
            } finally {
                applying = false;
            }
        }

        /** Faint highlight for the cell / header under the mouse, plus the right-click menu. */
        private void installHover() {
            table.addMouseMotionListener(new MouseAdapter() {
                @Override
                public void mouseMoved(MouseEvent e) {
                    if (drawing.isActive()) setHoverCell(-1, -1);
                    else setHoverCell(table.rowAtPoint(e.getPoint()), table.columnAtPoint(e.getPoint()));
                }
            });
            table.addMouseListener(new MouseAdapter() {
                @Override
                public void mouseExited(MouseEvent e) {
                    setHoverCell(-1, -1);
                }

                @Override
                public void mousePressed(MouseEvent e) {
                    if (e.isPopupTrigger()) showCellMenu(SheetView.this, e);
                }

                @Override
                public void mouseReleased(MouseEvent e) {
                    if (e.isPopupTrigger()) showCellMenu(SheetView.this, e);
                }
            });
            JTableHeader header = table.getTableHeader();
            header.addMouseMotionListener(new MouseAdapter() {
                @Override
                public void mouseMoved(MouseEvent e) {
                    int c = header.columnAtPoint(e.getPoint());
                    if (c != colHeaderRenderer.hover) {
                        colHeaderRenderer.hover = c;
                        header.repaint();
                    }
                }
            });
            header.addMouseListener(new MouseAdapter() {
                @Override
                public void mouseExited(MouseEvent e) {
                    colHeaderRenderer.hover = -1;
                    header.repaint();
                }
            });
            rowHeader.addMouseMotionListener(new MouseAdapter() {
                @Override
                public void mouseMoved(MouseEvent e) {
                    int i = rowHeader.locationToIndex(e.getPoint());
                    Rectangle rb = i < 0 ? null : rowHeader.getCellBounds(i, i);
                    if (rb == null || !rb.contains(e.getPoint())) i = -1;
                    if (i != rowHeaderRenderer.hover) {
                        rowHeaderRenderer.hover = i;
                        rowHeader.repaint();
                    }
                }
            });
            rowHeader.addMouseListener(new MouseAdapter() {
                @Override
                public void mouseExited(MouseEvent e) {
                    rowHeaderRenderer.hover = -1;
                    rowHeader.repaint();
                }
            });
        }

        private void setHoverCell(int r, int c) {
            if (r == cellRenderer.hoverRow && c == cellRenderer.hoverCol) return;
            repaintCell(cellRenderer.hoverRow, cellRenderer.hoverCol);
            cellRenderer.hoverRow = r;
            cellRenderer.hoverCol = c;
            repaintCell(r, c);
        }

        private void repaintCell(int r, int c) {
            if (r >= 0 && c >= 0) table.repaint(table.getCellRect(r, c, false));
        }

        private void installZoomAndResize() {
            // Remember widths the user drags (stored at 100% zoom so they survive zoom changes)
            table.getColumnModel().addColumnModelListener(new TableColumnModelListener() {
                @Override
                public void columnMarginChanged(javax.swing.event.ChangeEvent e) {
                    if (applying) return;
                    boolean changed = false;
                    for (int c = 0; c < model.getColumnCount(); c++) {
                        int w = table.getColumnModel().getColumn(c).getWidth();
                        if (w != scaled(model.getColumnWidth(c))) {
                            model.setColumnWidth(c, (int) Math.round(w / factor()));
                            changed = true;
                        }
                    }
                    if (changed) {
                        markDirty();
                        table.repaint();
                    }
                }

                @Override
                public void columnAdded(TableColumnModelEvent e) {
                }

                @Override
                public void columnRemoved(TableColumnModelEvent e) {
                }

                @Override
                public void columnMoved(TableColumnModelEvent e) {
                }

                @Override
                public void columnSelectionChanged(javax.swing.event.ListSelectionEvent e) {
                }
            });

            // Double-click a column border to fit the column to its content
            JTableHeader header = table.getTableHeader();
            header.addMouseListener(new MouseAdapter() {
                @Override
                public void mouseClicked(MouseEvent e) {
                    if (e.getClickCount() != 2) return;
                    int col = header.columnAtPoint(e.getPoint());
                    if (col < 0) return;
                    Rectangle r = header.getHeaderRect(col);
                    int target;
                    if (e.getX() - r.x <= 4) target = col - 1;
                    else if (r.x + r.width - e.getX() <= 4) target = col;
                    else return;
                    if (target >= 0) autoFitColumn(target);
                }
            });

            // Ctrl + mouse wheel zooms; the plain wheel keeps scrolling
            MouseWheelListener[] old = scroll.getMouseWheelListeners();
            for (MouseWheelListener l : old) scroll.removeMouseWheelListener(l);
            scroll.addMouseWheelListener(e -> {
                if (e.isControlDown()) {
                    zoomBy(e.getWheelRotation() < 0 ? 10 : -10);
                    e.consume();
                } else {
                    for (MouseWheelListener l : old) l.mouseWheelMoved(e);
                }
            });
        }

        void autoFitColumn(int c) {
            Font base = new Font(Font.SANS_SERIF, Font.PLAIN, BASE_FONT);
            int max = table.getFontMetrics(base.deriveFont(Font.BOLD)).stringWidth(FormulaEngine.columnLabel(c));
            for (int r = 0; r < model.getRowCount(); r++) {
                Object v = model.getValueAt(r, c);
                if (v == null || v.toString().isEmpty()) continue;
                CellStyle st = model.getStyle(r, c);
                int flags = (st != null && st.bold ? Font.BOLD : 0) | (st != null && st.italic ? Font.ITALIC : 0);
                float size = st != null && st.fontSize > 0 ? st.fontSize : BASE_FONT;
                FontMetrics fm = table.getFontMetrics(base.deriveFont(flags, size));
                max = Math.max(max, fm.stringWidth(v.toString()));
            }
            model.setColumnWidth(c, max + 16);
            applyZoom();
            markDirty();
            status("Column " + FormulaEngine.columnLabel(c) + " fitted to content");
        }

        private void installKeyBindings() {
            int ctrl = InputEvent.CTRL_DOWN_MASK;
            bind(KeyEvent.VK_Z, ctrl, "undoDrawing", () -> drawUndo());
            bind(KeyEvent.VK_Y, ctrl, "redoDrawing", () -> drawRedo());
            bind(KeyEvent.VK_ADD, ctrl, "zoomInPad", () -> zoomBy(10));
            bind(KeyEvent.VK_PLUS, ctrl, "zoomInPlus", () -> zoomBy(10));
            bind(KeyEvent.VK_SUBTRACT, ctrl, "zoomOutPad", () -> zoomBy(-10));
            bind(KeyEvent.VK_NUMPAD0, ctrl, "zoomResetPad", () -> setZoom(100));
            bind(KeyEvent.VK_B, InputEvent.CTRL_DOWN_MASK, "bold", () -> toggleStyle(true));
            bind(KeyEvent.VK_I, InputEvent.CTRL_DOWN_MASK, "italic", () -> toggleStyle(false));
            bind(KeyEvent.VK_L, InputEvent.CTRL_DOWN_MASK, "alignLeft", () -> align(SwingConstants.LEFT));
            bind(KeyEvent.VK_E, InputEvent.CTRL_DOWN_MASK, "alignCenter", () -> align(SwingConstants.CENTER));
            bind(KeyEvent.VK_R, InputEvent.CTRL_DOWN_MASK, "alignRight", () -> align(SwingConstants.RIGHT));
            bind(KeyEvent.VK_DELETE, 0, "clearContents", () -> {
                if (layer.hasSelection()) layer.deleteSelected();
                else clearSelection();
            });
        }

        private void bind(int key, int modifiers, String name, Runnable action) {
            table.getInputMap(JComponent.WHEN_FOCUSED).put(KeyStroke.getKeyStroke(key, modifiers), name);
            table.getActionMap().put(name, new AbstractAction() {
                @Override
                public void actionPerformed(ActionEvent e) {
                    action.run();
                }
            });
        }
    }

    /** In-cell editor that shows the raw input (formula text) instead of the computed value. */
    private static final class FormulaCellEditor extends DefaultCellEditor {
        private final SpreadsheetModel model;

        FormulaCellEditor(SpreadsheetModel model, Font font) {
            super(new JTextField());
            this.model = model;
            setClickCountToStart(2);
            JTextField field = (JTextField) getComponent();
            field.setFont(font);
            field.setBorder(BorderFactory.createLineBorder(GREEN, 2));
        }

        void setFont(Font font) {
            getComponent().setFont(font);
        }

        @Override
        public Component getTableCellEditorComponent(JTable table, Object value, boolean isSelected,
                                                     int row, int column) {
            return super.getTableCellEditorComponent(table, model.getInput(row, column), isSelected, row, column);
        }
    }

    // ==================================================================
    // Actions on the selection
    // ==================================================================

    private void applyToSelection(Consumer<CellStyle> change) {
        SheetView sv = currentSheet();
        if (sv == null || !guardEditable(sv)) return;
        int[] rows = sv.table.getSelectedRows();
        int[] cols = sv.table.getSelectedColumns();
        sv.model.beginBatch();
        try {
            for (int r : rows) for (int c : cols) sv.model.updateStyle(r, c, change);
        } finally {
            sv.model.endBatch();
        }
        sv.table.requestFocusInWindow();
    }

    /** Excel behaviour: if every selected cell already has the style, remove it; otherwise apply it. */
    private void toggleStyle(boolean bold) {
        SheetView sv = currentSheet();
        if (sv == null) return;
        int[] rows = sv.table.getSelectedRows();
        int[] cols = sv.table.getSelectedColumns();
        if (rows.length == 0 || cols.length == 0) return;
        boolean all = true;
        outer:
        for (int r : rows) {
            for (int c : cols) {
                CellStyle st = sv.model.getStyle(r, c);
                boolean on = st != null && (bold ? st.bold : st.italic);
                if (!on) {
                    all = false;
                    break outer;
                }
            }
        }
        final boolean value = !all;
        applyToSelection(st -> {
            if (bold) st.bold = value;
            else st.italic = value;
        });
    }

    private void align(int alignment) {
        applyToSelection(st -> st.alignment = alignment);
    }

    private void pickFillColor() {
        if (!guardEditable(currentSheet())) return;
        Color c = javax.swing.JColorChooser.showDialog(this, "Choose Background Color", currentBackground);
        if (c == null) return;
        currentBackground = c;
        fillChip.setBackground(c);
        applyToSelection(st -> st.background = c);
    }

    private void clearSelection() {
        SheetView sv = currentSheet();
        if (sv == null || !guardEditable(sv)) return;
        int[] rows = sv.table.getSelectedRows();
        int[] cols = sv.table.getSelectedColumns();
        sv.model.beginBatch();
        try {
            for (int r : rows) for (int c : cols) sv.model.clearContents(r, c);
        } finally {
            sv.model.endBatch();
        }
    }

    /**
     * AutoSum. Single cell: sums the contiguous numbers above (or, failing that, to the left).
     * Multi-cell selection: puts a SUM for each selected column in the row below the selection.
     */
    private void autoSum() {
        SheetView sv = currentSheet();
        if (sv == null || !guardEditable(sv)) return;
        SpreadsheetModel m = sv.model;
        int[] rows = sv.table.getSelectedRows();
        int[] cols = sv.table.getSelectedColumns();
        if (rows.length == 0 || cols.length == 0) {
            status("Select a cell first");
            return;
        }
        int r1 = rows[0], r2 = rows[rows.length - 1], c1 = cols[0], c2 = cols[cols.length - 1];

        if (r1 == r2 && c1 == c2) {
            int top = r1;
            while (top - 1 >= 0 && m.isNumeric(top - 1, c1)) top--;
            if (top < r1) {
                m.setInput(r1, c1, "=SUM(" + FormulaEngine.cellRef(top, c1) + ":" + FormulaEngine.cellRef(r1 - 1, c1) + ")");
                return;
            }
            int left = c1;
            while (left - 1 >= 0 && m.isNumeric(r1, left - 1)) left--;
            if (left < c1) {
                m.setInput(r1, c1, "=SUM(" + FormulaEngine.cellRef(r1, left) + ":" + FormulaEngine.cellRef(r1, c1 - 1) + ")");
                return;
            }
            m.setInput(r1, c1, "=SUM()");
            status("No adjacent numbers found - type the range inside SUM( )");
            formulaField.requestFocusInWindow();
            formulaField.setCaretPosition(formulaField.getText().length() - 1);
        } else {
            int target = r2 + 1;
            if (target >= m.getRowCount()) {
                status("No room below the selection for the totals");
                return;
            }
            m.beginBatch();
            try {
                for (int c : cols) {
                    m.setInput(target, c, "=SUM(" + FormulaEngine.cellRef(r1, c) + ":" + FormulaEngine.cellRef(r2, c) + ")");
                }
            } finally {
                m.endBatch();
            }
        }
    }

    // ==================================================================
    // Dialogs
    // ==================================================================

    private void openInsertDialog() {
        SheetView sv = currentSheet();
        if (sv == null || !guardEditable(sv)) return;
        int[] rows = sv.table.getSelectedRows();
        int[] cols = sv.table.getSelectedColumns();
        int r = rows.length > 0 ? rows[0] : 0;
        int c = cols.length > 0 ? cols[0] : 0;
        new InsertDialog(this, sv.model, r, c, () -> {
            syncFormulaBar();
            sv.table.requestFocusInWindow();
            status("Inserted at " + FormulaEngine.cellRef(r, c));
        }).setVisible(true);
    }

    private void openMacroIde() {
        if (macroDialog == null) {
            macroDialog = new MacroIDEDialog(this, () -> {
                SheetView sv = currentSheet();
                return sv == null ? null : sv.model;
            });
        }
        macroDialog.setVisible(true);
        macroDialog.toFront();
    }

    // ==================================================================
    // CSV import / export
    // ==================================================================

    // ==================================================================
    // Sheet protection (Review menu)
    // ==================================================================

    /** @return false (with a hint in the status bar) when the sheet is protected. */
    private boolean guardEditable(SheetView sv) {
        if (sv != null && sv.model.isProtected()) {
            status("This sheet is protected. Use Review > Unprotect Sheet to make changes.");
            Toolkit.getDefaultToolkit().beep();
            return false;
        }
        return true;
    }

    private void updateProtectUi(SheetView sv) {
        boolean prot = sv != null && sv.model.isProtected();
        if (protectButton != null) {
            protectButton.setText(prot ? "Unprotect" : "Protect Sheet");
            protectButton.setIconName(prot ? "unlock" : "lock");
        }
        protLabel.setText(prot ? "SHEET PROTECTED" : "");
        formulaField.setEditable(!prot);
    }

    private void refreshTabLocks() {
        for (int i = 0; i < tabs.getTabCount(); i++) {
            Component comp = tabs.getComponentAt(i);
            Component header = tabs.getTabComponentAt(i);
            if (comp instanceof SheetView && header instanceof TabHeader) {
                ((TabHeader) header).setLocked(((SheetView) comp).model.isProtected());
            }
        }
    }

    private void toggleProtection() {
        SheetView sv = currentSheet();
        if (sv == null) return;
        if (sv.model.isProtected()) unprotectSheetDialog();
        else protectSheetDialog();
    }

    private void protectSheetDialog() {
        SheetView sv = currentSheet();
        if (sv == null) return;
        if (sv.model.isProtected()) {
            JOptionPane.showMessageDialog(this, "This sheet is already protected.", "Protect Sheet",
                    JOptionPane.INFORMATION_MESSAGE);
            return;
        }
        char[] pw = askNewPassword("Protect Sheet",
                "Password (optional - leave blank to protect without a password):",
                "Protected sheets cannot be edited, formatted, cleared, changed by macros, or have pictures moved.");
        if (pw == null) return;
        if (sv.table.isEditing()) sv.table.getCellEditor().cancelCellEditing();
        sv.model.getProtection().protect(pw);
        Arrays.fill(pw, '\0');
        afterProtectionChange(sv, "Sheet protected");
    }

    private void unprotectSheetDialog() {
        SheetView sv = currentSheet();
        if (sv == null || !sv.model.isProtected()) return;
        char[] pw = null;
        if (sv.model.getProtection().hasPassword()) {
            pw = askPassword("Unprotect Sheet", "Password for \"" + tabs.getTitleAt(tabs.getSelectedIndex()) + "\":");
            if (pw == null) return;
        }
        boolean ok = sv.model.getProtection().unprotect(pw);
        if (pw != null) Arrays.fill(pw, '\0');
        if (!ok) {
            JOptionPane.showMessageDialog(this, "Incorrect password.", "Unprotect Sheet", JOptionPane.ERROR_MESSAGE);
            return;
        }
        afterProtectionChange(sv, "Sheet unprotected");
    }

    private void afterProtectionChange(SheetView sv, String message) {
        refreshTabLocks();
        markDirty();
        sv.table.repaint();
        syncFormulaBar();
        status(message);
        sv.table.requestFocusInWindow();
    }

    // ---- password prompts ----

    /** Asks for a password once. @return the characters, or null if cancelled. */
    private char[] askPassword(String title, String message) {
        JPasswordField field = new JPasswordField(18);
        focusOnShow(field);
        JPanel panel = new JPanel(new GridLayout(0, 1, 4, 4));
        panel.add(new JLabel(message));
        panel.add(field);
        int answer = JOptionPane.showConfirmDialog(this, panel, title,
                JOptionPane.OK_CANCEL_OPTION, JOptionPane.PLAIN_MESSAGE);
        return answer == JOptionPane.OK_OPTION ? field.getPassword() : null;
    }

    /**
     * Asks for a new password with confirmation.
     * @return the characters (possibly empty), or null if cancelled or the two entries differ
     */
    private char[] askNewPassword(String title, String message, String note) {
        JPasswordField first = new JPasswordField(18);
        JPasswordField second = new JPasswordField(18);
        focusOnShow(first);
        JPanel panel = new JPanel(new GridLayout(0, 1, 4, 4));
        panel.add(new JLabel(message));
        panel.add(first);
        panel.add(new JLabel("Confirm password:"));
        panel.add(second);
        panel.add(new JLabel("<html><small>" + note + "</small></html>"));
        int answer = JOptionPane.showConfirmDialog(this, panel, title,
                JOptionPane.OK_CANCEL_OPTION, JOptionPane.PLAIN_MESSAGE);
        if (answer != JOptionPane.OK_OPTION) return null;
        char[] a = first.getPassword();
        char[] b = second.getPassword();
        boolean same = Arrays.equals(a, b);
        Arrays.fill(b, '\0');
        if (!same) {
            Arrays.fill(a, '\0');
            JOptionPane.showMessageDialog(this, "The passwords do not match.", title, JOptionPane.ERROR_MESSAGE);
            return null;
        }
        return a;
    }

    /** Puts the keyboard cursor in the field when its dialog appears. */
    private static void focusOnShow(JComponent field) {
        field.addAncestorListener(new AncestorListener() {
            @Override
            public void ancestorAdded(AncestorEvent event) {
                SwingUtilities.invokeLater(field::requestFocusInWindow);
            }

            @Override
            public void ancestorRemoved(AncestorEvent event) {
            }

            @Override
            public void ancestorMoved(AncestorEvent event) {
            }
        });
    }

    /** Small padlock drawn on the tab of a protected sheet. */
    private static final class LockIcon implements Icon {
        @Override
        public void paintIcon(Component c, Graphics g, int x, int y) {
            Graphics2D g2 = (Graphics2D) g.create();
            g2.setRenderingHint(RenderingHints.KEY_ANTIALIASING, RenderingHints.VALUE_ANTIALIAS_ON);
            g2.translate(x, y);
            g2.setColor(new Color(0x8A6D00));
            g2.setStroke(new BasicStroke(1.6f));
            g2.drawArc(2, 0, 7, 8, 0, 180);
            g2.drawLine(2, 4, 2, 6);
            g2.drawLine(9, 4, 9, 6);
            g2.setColor(new Color(0xD9A900));
            g2.fillRoundRect(0, 6, 11, 7, 3, 3);
            g2.dispose();
        }

        @Override
        public int getIconWidth() {
            return 11;
        }

        @Override
        public int getIconHeight() {
            return 13;
        }
    }

    // ==================================================================
    // Pictures
    // ==================================================================

    private void insertPictureFromFile() {
        SheetView sv = currentSheet();
        if (sv == null || !guardEditable(sv)) return;
        JFileChooser fc = new JFileChooser();
        fc.setFileFilter(new FileNameExtensionFilter("Pictures (png, jpg, gif, bmp)",
                "png", "jpg", "jpeg", "gif", "bmp"));
        if (fc.showOpenDialog(this) != JFileChooser.APPROVE_OPTION) return;
        try {
            BufferedImage img = ImageIO.read(fc.getSelectedFile());
            if (img == null) {
                JOptionPane.showMessageDialog(this, "That file is not a supported picture (use PNG, JPG, GIF or BMP).",
                        "Insert Picture", JOptionPane.ERROR_MESSAGE);
                return;
            }
            placePicture(sv, img);
        } catch (IOException | RuntimeException ex) {
            JOptionPane.showMessageDialog(this, "Could not read the picture:\n" + ex.getMessage(),
                    "Insert Picture", JOptionPane.ERROR_MESSAGE);
        }
    }

    private void insertPictureFromClipboard() {
        SheetView sv = currentSheet();
        if (sv == null || !guardEditable(sv)) return;
        try {
            Clipboard cb = Toolkit.getDefaultToolkit().getSystemClipboard();
            if (!cb.isDataFlavorAvailable(DataFlavor.imageFlavor)) {
                JOptionPane.showMessageDialog(this, "There is no picture on the clipboard.",
                        "Insert Picture", JOptionPane.INFORMATION_MESSAGE);
                return;
            }
            java.awt.Image im = (java.awt.Image) cb.getData(DataFlavor.imageFlavor);
            javax.swing.ImageIcon icon = new javax.swing.ImageIcon(im); // makes sure it is fully loaded
            int w = icon.getIconWidth();
            int h = icon.getIconHeight();
            if (w <= 0 || h <= 0) throw new IOException("The clipboard picture is empty");
            BufferedImage bi = new BufferedImage(w, h, BufferedImage.TYPE_INT_ARGB);
            Graphics2D g = bi.createGraphics();
            g.drawImage(im, 0, 0, null);
            g.dispose();
            placePicture(sv, bi);
        } catch (java.awt.datatransfer.UnsupportedFlavorException | IOException | RuntimeException ex) {
            JOptionPane.showMessageDialog(this, "Could not paste the picture:\n" + ex.getMessage(),
                    "Insert Picture", JOptionPane.ERROR_MESSAGE);
        }
    }

    /** Puts the picture at the active cell and selects it so it can be moved or resized. */
    private void placePicture(SheetView sv, BufferedImage img) throws IOException {
        int r = Math.max(0, sv.table.getSelectionModel().getLeadSelectionIndex());
        int c = Math.max(0, sv.table.getColumnModel().getSelectionModel().getLeadSelectionIndex());
        Rectangle cell = sv.table.getCellRect(r, c, true);
        double zoom = sv.model.getZoomPercent() / 100.0; // pictures are stored in 100%-zoom units
        SheetImage si = SheetImage.create(img, (int) Math.round(cell.x / zoom), (int) Math.round(cell.y / zoom));
        sv.model.addImage(si);
        sv.layer.select(si);
        markDirty();
        sv.table.requestFocusInWindow();
        status("Picture inserted at " + FormulaEngine.cellRef(r, c)
                + " - drag to move, drag the corner square to resize, Delete to remove");
    }

    private void deleteSelectedPicture() {
        SheetView sv = currentSheet();
        if (sv == null || !guardEditable(sv)) return;
        if (sv.layer.deleteSelected()) status("Picture deleted");
        else status("Click a picture to select it first");
    }

    // ==================================================================
    // Workbook files (.spx), open password, unsaved-changes tracking
    // ==================================================================

    private void markDirty() {
        if (loading || dirty) return;
        dirty = true;
        updateTitle();
    }

    private void updateTitle() {
        String name = currentFile == null ? "Untitled" : currentFile.getName();
        setTitle("Spreadsheet Pro - " + name + (dirty ? " *" : ""));
    }

    /** @return true if it is fine to replace the current workbook (saved, discarded or unchanged). */
    private boolean confirmDiscard() {
        if (!dirty) return true;
        String name = currentFile == null ? "Untitled" : currentFile.getName();
        int answer = JOptionPane.showConfirmDialog(this, "Save changes to " + name + "?", "Unsaved Changes",
                JOptionPane.YES_NO_CANCEL_OPTION, JOptionPane.WARNING_MESSAGE);
        if (answer == JOptionPane.YES_OPTION) return saveWorkbook();
        return answer == JOptionPane.NO_OPTION;
    }

    private void exitApp() {
        if (!confirmDiscard()) return;
        dispose();
        System.exit(0);
    }

    private void newWorkbook() {
        if (!confirmDiscard()) return;
        loading = true;
        try {
            tabs.removeAll();
            sheetCounter = 0;
            addSheet(nextSheetName());
        } finally {
            loading = false;
        }
        currentFile = null;
        if (workbookPassword != null) Arrays.fill(workbookPassword, '\0');
        workbookPassword = null;
        dirty = false;
        updateTitle();
        refreshTabLocks();
        syncFormulaBar();
        status("New workbook");
    }

    private JFileChooser workbookChooser() {
        JFileChooser fc = new JFileChooser(currentFile != null ? currentFile.getParentFile() : null);
        fc.setFileFilter(new FileNameExtensionFilter("Spreadsheet Pro workbook (*.spx)", WorkbookIO.EXTENSION));
        return fc;
    }

    private List<WorkbookIO.SheetData> collectSheets() {
        List<WorkbookIO.SheetData> list = new ArrayList<>();
        for (int i = 0; i < tabs.getTabCount(); i++) {
            Component comp = tabs.getComponentAt(i);
            if (comp instanceof SheetView) {
                list.add(new WorkbookIO.SheetData(tabs.getTitleAt(i), ((SheetView) comp).model));
            }
        }
        return list;
    }

    /** Ctrl+S. @return true if the workbook was saved. */
    private boolean saveWorkbook() {
        if (currentFile == null) return saveWorkbookAs();
        return writeWorkbook(currentFile);
    }

    private boolean saveWorkbookAs() {
        JFileChooser fc = workbookChooser();
        fc.setSelectedFile(new File(currentFile != null ? currentFile.getName() : "Workbook.spx"));
        if (fc.showSaveDialog(this) != JFileChooser.APPROVE_OPTION) return false;
        File f = fc.getSelectedFile();
        if (!f.getName().toLowerCase().endsWith("." + WorkbookIO.EXTENSION)) {
            f = new File(f.getParentFile(), f.getName() + "." + WorkbookIO.EXTENSION);
        }
        if (f.exists() && JOptionPane.showConfirmDialog(this, "\"" + f.getName() + "\" already exists. Replace it?",
                "Save As", JOptionPane.YES_NO_OPTION, JOptionPane.WARNING_MESSAGE) != JOptionPane.YES_OPTION) {
            return false;
        }
        if (!writeWorkbook(f)) return false;
        currentFile = f;
        updateTitle();
        return true;
    }

    private boolean writeWorkbook(File f) {
        SheetView sv = currentSheet();
        if (sv != null && sv.table.isEditing()) sv.table.getCellEditor().stopCellEditing();
        try {
            WorkbookIO.save(f, collectSheets(), tabs.getSelectedIndex(), workbookPassword);
            dirty = false;
            updateTitle();
            status("Saved " + f.getName() + (workbookPassword != null ? " (encrypted)" : ""));
            return true;
        } catch (IOException ex) {
            JOptionPane.showMessageDialog(this, "Could not save the workbook:\n" + ex.getMessage(),
                    "Save", JOptionPane.ERROR_MESSAGE);
            return false;
        }
    }

    private void openWorkbook() {
        if (!confirmDiscard()) return;
        JFileChooser fc = workbookChooser();
        if (fc.showOpenDialog(this) != JFileChooser.APPROVE_OPTION) return;
        File f = fc.getSelectedFile();
        char[] pw = null;
        WorkbookIO.WorkbookData data;
        while (true) {
            try {
                data = WorkbookIO.load(f, pw);
                break;
            } catch (WorkbookIO.PasswordRequiredException | WorkbookIO.BadPasswordException ex) {
                boolean wrong = ex instanceof WorkbookIO.BadPasswordException;
                pw = askPassword("Open Workbook",
                        wrong ? "Incorrect password. Try again:" : "\"" + f.getName() + "\" needs a password to open:");
                if (pw == null) return;
            } catch (IOException ex) {
                JOptionPane.showMessageDialog(this, "Could not open the workbook:\n" + ex.getMessage(),
                        "Open", JOptionPane.ERROR_MESSAGE);
                return;
            }
        }

        loading = true;
        try {
            tabs.removeAll();
            for (WorkbookIO.SheetData sd : data.sheets) addSheet(sd.name, sd.model);
            tabs.setSelectedIndex(data.activeIndex);
        } finally {
            loading = false;
        }
        sheetCounter = tabs.getTabCount();
        currentFile = f;
        if (workbookPassword != null) Arrays.fill(workbookPassword, '\0');
        workbookPassword = pw;
        dirty = false;
        updateTitle();
        refreshTabLocks();
        syncFormulaBar();
        status("Opened " + f.getName());
    }

    private void setOpenPassword() {
        char[] pw = askNewPassword("Set Open Password",
                "Password needed to open this workbook (leave blank to remove it):",
                "The workbook is encrypted when you save it. If you forget the password the file cannot be recovered.");
        if (pw == null) return;
        if (workbookPassword != null) Arrays.fill(workbookPassword, '\0');
        workbookPassword = pw.length == 0 ? null : pw;
        markDirty();
        status(pw.length == 0 ? "Open password removed - save to apply"
                : "Open password set - save the workbook to encrypt it");
    }

    private void exportCsv() {
        SheetView sv = currentSheet();
        if (sv == null) return;
        JFileChooser fc = new JFileChooser();
        fc.setSelectedFile(new File(tabs.getTitleAt(tabs.getSelectedIndex()) + ".csv"));
        if (fc.showSaveDialog(this) != JFileChooser.APPROVE_OPTION) return;
        File file = fc.getSelectedFile();
        try {
            Files.write(file.toPath(), toCsv(sv.model).getBytes(StandardCharsets.UTF_8));
            status("Exported " + file.getName());
        } catch (IOException ex) {
            JOptionPane.showMessageDialog(this, "Could not write file:\n" + ex.getMessage(),
                    "Export CSV", JOptionPane.ERROR_MESSAGE);
        }
    }

    private void importCsv() {
        JFileChooser fc = new JFileChooser();
        if (fc.showOpenDialog(this) != JFileChooser.APPROVE_OPTION) return;
        File file = fc.getSelectedFile();
        try {
            String content = new String(Files.readAllBytes(file.toPath()), StandardCharsets.UTF_8);
            String name = file.getName().replaceFirst("\\.[^.]*$", "");
            SheetView sv = addSheet(name.isEmpty() ? nextSheetName() : name);
            List<List<String>> rows = parseCsv(content);
            sv.model.beginBatch();
            try {
                for (int r = 0; r < rows.size(); r++) {
                    List<String> row = rows.get(r);
                    for (int c = 0; c < row.size(); c++) sv.model.setInput(r, c, row.get(c));
                }
            } finally {
                sv.model.endBatch();
            }
            status("Imported " + file.getName());
        } catch (IOException ex) {
            JOptionPane.showMessageDialog(this, "Could not read file:\n" + ex.getMessage(),
                    "Import CSV", JOptionPane.ERROR_MESSAGE);
        }
    }

    private static String toCsv(SpreadsheetModel m) {
        StringBuilder sb = new StringBuilder();
        int rows = m.getUsedRowCount();
        int cols = m.getUsedColumnCount();
        for (int r = 0; r < rows; r++) {
            for (int c = 0; c < cols; c++) {
                if (c > 0) sb.append(',');
                String v = m.getInput(r, c);
                if (v.contains(",") || v.contains("\"") || v.contains("\n")) {
                    v = "\"" + v.replace("\"", "\"\"") + "\"";
                }
                sb.append(v);
            }
            sb.append("\r\n");
        }
        return sb.toString();
    }

    static List<List<String>> parseCsv(String s) {
        List<List<String>> rows = new ArrayList<>();
        List<String> row = new ArrayList<>();
        StringBuilder cell = new StringBuilder();
        boolean inQuotes = false;
        for (int i = 0; i < s.length(); i++) {
            char ch = s.charAt(i);
            if (inQuotes) {
                if (ch == '"') {
                    if (i + 1 < s.length() && s.charAt(i + 1) == '"') {
                        cell.append('"');
                        i++;
                    } else {
                        inQuotes = false;
                    }
                } else {
                    cell.append(ch);
                }
            } else if (ch == '"') {
                inQuotes = true;
            } else if (ch == ',') {
                row.add(cell.toString());
                cell.setLength(0);
            } else if (ch == '\n') {
                row.add(cell.toString());
                cell.setLength(0);
                rows.add(row);
                row = new ArrayList<>();
            } else if (ch != '\r') {
                cell.append(ch);
            }
        }
        if (cell.length() > 0 || !row.isEmpty()) {
            row.add(cell.toString());
            rows.add(row);
        }
        return rows;
    }
}

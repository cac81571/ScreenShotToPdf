package com.image2pdf;

import com.formdev.flatlaf.FlatLightLaf;
import com.lowagie.text.Chunk;
import com.lowagie.text.Document;
import com.lowagie.text.Font;
import com.lowagie.text.Image;
import com.lowagie.text.PageSize;
import com.lowagie.text.Paragraph;
import com.lowagie.text.pdf.BaseFont;
import com.lowagie.text.pdf.PdfWriter;

import javax.swing.*;
import javax.swing.border.EmptyBorder;
import javax.swing.table.DefaultTableModel;
import javax.swing.table.TableColumnModel;
import java.awt.*;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardCopyOption;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Comparator;
import java.util.HashSet;
import java.util.List;
import java.util.Locale;
import java.util.Set;
import java.util.stream.Collectors;
import java.util.stream.Stream;

/**
 * システムマイグレーションのスクリーンショットを
 * 移行前・移行後で比較できるPDFにまとめるSwingアプリケーション。
 * 奇数ページ＝移行前、偶数ページ＝移行後で交互に配置する。
 */
public class MigrationScreenshotPdfApp {

    /** 空白ページとしてリストに表示するラベル（UI用・PDF用とも日本語） */
    private static final String BLANK_PAGE_LABEL = "(空白ページ)";
    /** 画像として扱うファイル拡張子の一覧 */
    private static final List<String> IMAGE_EXTENSIONS = Arrays.asList(
            ".png", ".jpg", ".jpeg", ".gif", ".bmp", ".webp");
    /** フォルダ履歴に保持する最大件数 */
    private static final int HISTORY_MAX = 30;
    /** 差分未計算・比較不可行の行に表示する記号 */
    private static final String DIFF_PLACEHOLDER = "—";
    /** 履歴ファイルを保存するディレクトリ（ユーザーホーム直下の .screenShotToPdf） */
    private static final Path HISTORY_DIR = Paths.get(
            System.getProperty("user.home"), ".screenShotToPdf");

    /** PDF用テキストフォント（日本語対応を試行し、失敗時は Helvetica）。null の場合は都度取得。 */
    private static Font pdfTextFont;

    /** 画像フォルダAのパス入力・選択用コンボボックス */
    private JComboBox<String> beforeFolderCombo;
    /** 画像フォルダBのパス入力・選択用コンボボックス */
    private JComboBox<String> afterFolderCombo;
    /** PDF作成を実行するボタン */
    private JButton createPdfButton;
    /** 選択行の A/B 画像に差分マスクを重ねて PNG 保存するボタン */
    private JButton diffOverlayButton;
    /** ログメッセージを表示するテキストエリア */
    private JTextArea logArea;
    /** 画像ファイルA（# / 差分 / ファイル名） */
    private DefaultTableModel beforeTableModel;
    /** 画像ファイルB（# / 差分 / ファイル名） */
    private DefaultTableModel afterTableModel;
    private JTable beforeTable;
    private JTable afterTable;
    /** 画像ファイルAのPDF表示タイトル入力欄 */
    private JTextField beforeTitleField;
    /** 画像ファイルBのPDF表示タイトル入力欄 */
    private JTextField afterTitleField;
    /** 差分判定しきい値入力欄（RGB差分合計） */
    private JTextField diffThresholdField;
    /** 差分マスク赤アルファ入力欄（0.0〜1.0） */
    private JTextField diffAlphaField;
    /** 差分マスク拡張半径入力欄（px） */
    private JTextField diffExpandRadiusField;
    /** 差分PNGの画像Aに赤マスクを重ねるか */
    private JCheckBox diffOverlayCheckA;
    /** 差分PNGの画像Bに赤マスクを重ねるか */
    private JCheckBox diffOverlayCheckB;
    /** 画像ファイルAリストに対応するファイルパス（null は空白ページ） */
    private List<Path> beforeFilePaths = new ArrayList<>();
    /** 画像ファイルBリストに対応するファイルパス（null は空白ページ） */
    private List<Path> afterFilePaths = new ArrayList<>();

    /**
     * エントリポイント。FlatLaf を設定し、Swing の EDT で GUI を起動する。
     *
     * @param args コマンドライン引数（未使用）
     */
    public static void main(String[] args) {
        FlatLightLaf.setup();
        SwingUtilities.invokeLater(() -> new MigrationScreenshotPdfApp().run());
    }

    /**
     * メインウィンドウを構築し、フォルダ選択・リスト・PDF作成ボタン・ログ領域を配置して表示する。
     */
    void run() {
        JFrame frame = new JFrame("マイグレーション スクリーンショット PDF 作成ツール");
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        frame.setSize(1040, 560);
        frame.setLocationRelativeTo(null);

        JPanel main = new JPanel(new BorderLayout(10, 10));
        main.setBorder(new EmptyBorder(15, 15, 15, 15));

        JPanel form = new JPanel();
        form.setLayout(new BoxLayout(form, BoxLayout.PAGE_AXIS));
        form.setAlignmentX(Component.LEFT_ALIGNMENT);

        beforeFolderCombo = new JComboBox<>(loadHistory("before_folders").toArray(new String[0]));
        afterFolderCombo = new JComboBox<>(loadHistory("after_folders").toArray(new String[0]));
        for (JComboBox<String> cb : Arrays.asList(beforeFolderCombo, afterFolderCombo)) {
            cb.setEditable(true);
            int h = (int) cb.getPreferredSize().getHeight();
            cb.setMaximumSize(new Dimension(Integer.MAX_VALUE, h));
        }

        JButton extractBtn = new JButton("ファイル抽出");
        setFixedButtonSize(extractBtn, 100, 35);
        extractBtn.addActionListener(e -> extractImageFilesFromBoth());

        form.add(createFolderRow("画像フォルダA:", beforeFolderCombo, true));
        form.add(Box.createVerticalStrut(8));
        form.add(createFolderRow("画像フォルダB:", afterFolderCombo, false));
        form.add(Box.createVerticalStrut(8));
        JPanel extractRow = new JPanel(new BorderLayout());
        extractRow.add(extractBtn, BorderLayout.EAST);
        form.add(extractRow);

        main.add(form, BorderLayout.NORTH);

        beforeTableModel = createFileTableModel();
        afterTableModel = createFileTableModel();
        beforeTable = new JTable(beforeTableModel);
        afterTable = new JTable(afterTableModel);
        configureFileTable(beforeTable);
        configureFileTable(afterTable);
        beforeTitleField = new JTextField("移行前", 10);
        afterTitleField = new JTextField("移行後", 10);

        JPanel listPanel = new JPanel(new GridLayout(1, 2, 10, 0));
        JPanel beforeListPanel = new JPanel(new BorderLayout(2, 4));
        beforeListPanel.setBorder(new EmptyBorder(2, 4, 2, 4));
        beforeListPanel.add(createTitleInputPanel("画像ファイルA", beforeTitleField), BorderLayout.NORTH);
        beforeListPanel.add(new JScrollPane(beforeTable), BorderLayout.CENTER);
        JPanel beforeBtnPanel = new JPanel(new FlowLayout(FlowLayout.LEFT, 4, 2));
        JButton beforeDeleteBtn = new JButton("選択削除");
        beforeDeleteBtn.addActionListener(e -> removeSelectedFromList(true));
        JButton beforeBlankBeforeBtn = new JButton("前に空白追加");
        beforeBlankBeforeBtn.addActionListener(e -> addBlankPageToList(true, true));
        JButton beforeBlankAfterBtn = new JButton("後ろに空白追加");
        beforeBlankAfterBtn.addActionListener(e -> addBlankPageToList(true, false));
        JButton beforeMoveUpBtn = new JButton("上へ");
        beforeMoveUpBtn.addActionListener(e -> moveSelectedRowInFileGrid(true, -1));
        JButton beforeMoveDownBtn = new JButton("下へ");
        beforeMoveDownBtn.addActionListener(e -> moveSelectedRowInFileGrid(true, 1));
        int beforeRowBtnH = beforeDeleteBtn.getPreferredSize().height;
        setFixedButtonSize(beforeMoveUpBtn, 52, beforeRowBtnH);
        setFixedButtonSize(beforeMoveDownBtn, 52, beforeRowBtnH);
        beforeBtnPanel.add(beforeDeleteBtn);
        beforeBtnPanel.add(beforeBlankBeforeBtn);
        beforeBtnPanel.add(beforeBlankAfterBtn);
        beforeBtnPanel.add(beforeMoveUpBtn);
        beforeBtnPanel.add(beforeMoveDownBtn);
        beforeListPanel.add(beforeBtnPanel, BorderLayout.SOUTH);

        JPanel afterListPanel = new JPanel(new BorderLayout(2, 4));
        afterListPanel.setBorder(new EmptyBorder(2, 4, 2, 4));
        afterListPanel.add(createTitleInputPanel("画像ファイルB", afterTitleField), BorderLayout.NORTH);
        afterListPanel.add(new JScrollPane(afterTable), BorderLayout.CENTER);
        JPanel afterBtnPanel = new JPanel(new FlowLayout(FlowLayout.LEFT, 4, 2));
        JButton afterDeleteBtn = new JButton("選択削除");
        afterDeleteBtn.addActionListener(e -> removeSelectedFromList(false));
        JButton afterBlankBeforeBtn = new JButton("前に空白追加");
        afterBlankBeforeBtn.addActionListener(e -> addBlankPageToList(false, true));
        JButton afterBlankAfterBtn = new JButton("後ろに空白追加");
        afterBlankAfterBtn.addActionListener(e -> addBlankPageToList(false, false));
        JButton afterMoveUpBtn = new JButton("上へ");
        afterMoveUpBtn.addActionListener(e -> moveSelectedRowInFileGrid(false, -1));
        JButton afterMoveDownBtn = new JButton("下へ");
        afterMoveDownBtn.addActionListener(e -> moveSelectedRowInFileGrid(false, 1));
        int afterRowBtnH = afterDeleteBtn.getPreferredSize().height;
        setFixedButtonSize(afterMoveUpBtn, 52, afterRowBtnH);
        setFixedButtonSize(afterMoveDownBtn, 52, afterRowBtnH);
        afterBtnPanel.add(afterDeleteBtn);
        afterBtnPanel.add(afterBlankBeforeBtn);
        afterBtnPanel.add(afterBlankAfterBtn);
        afterBtnPanel.add(afterMoveUpBtn);
        afterBtnPanel.add(afterMoveDownBtn);
        afterListPanel.add(afterBtnPanel, BorderLayout.SOUTH);
        listPanel.add(beforeListPanel);
        listPanel.add(afterListPanel);
        main.add(listPanel, BorderLayout.CENTER);

        JPanel bottom = new JPanel(new BorderLayout(0, 6));
        JPanel buttonRow = new JPanel(new BorderLayout(8, 0));
        buttonRow.setAlignmentX(Component.LEFT_ALIGNMENT);
        buttonRow.add(createDiffSettingsPanel(), BorderLayout.WEST);
        JPanel rightButtons = new JPanel(new FlowLayout(FlowLayout.RIGHT, 8, 0));
        diffOverlayButton = new JButton("差分PDF作成");
        setFixedButtonSize(diffOverlayButton, 120, 35);
        diffOverlayButton.addActionListener(e -> createDiffOverlayPdf());
        createPdfButton = new JButton("PDF作成");
        setFixedButtonSize(createPdfButton, 100, 35);
        createPdfButton.addActionListener(e -> createPdf());
        rightButtons.add(diffOverlayButton);
        rightButtons.add(createPdfButton);
        buttonRow.add(rightButtons, BorderLayout.EAST);
        bottom.add(buttonRow, BorderLayout.NORTH);

        logArea = new JTextArea(6, 40);
        logArea.setEditable(false);
        logArea.setLineWrap(true);
        logArea.setWrapStyleWord(true);
        bottom.add(new JScrollPane(logArea), BorderLayout.SOUTH);
        main.add(bottom, BorderLayout.SOUTH);

        frame.getContentPane().add(main);
        frame.setVisible(true);
    }

    /**
     * 差分画像保存で使うパラメータ入力欄を作成する。
     */
    private JPanel createDiffSettingsPanel() {
        JPanel panel = new JPanel(new FlowLayout(FlowLayout.LEFT, 6, 0));
        diffThresholdField = new JTextField("32", 4);
        diffAlphaField = new JTextField("0.45", 4);
        diffExpandRadiusField = new JTextField("10", 3);
        diffOverlayCheckA = new JCheckBox("Aにマスク", false);
        diffOverlayCheckB = new JCheckBox("Bにマスク", true);
        panel.add(new JLabel("差分しきい値:"));
        panel.add(diffThresholdField);
        panel.add(new JLabel("赤α:"));
        panel.add(diffAlphaField);
        panel.add(new JLabel("マスク拡張:"));
        panel.add(diffExpandRadiusField);
        panel.add(diffOverlayCheckA);
        panel.add(diffOverlayCheckB);
        return panel;
    }

    private static DefaultTableModel createFileTableModel() {
        return new DefaultTableModel(new Object[]{"#", "差分", "ファイル名"}, 0) {
            @Override
            public boolean isCellEditable(int row, int column) {
                return false;
            }

            @Override
            public Class<?> getColumnClass(int columnIndex) {
                return String.class;
            }
        };
    }

    private static void configureFileTable(JTable table) {
        table.setSelectionMode(ListSelectionModel.MULTIPLE_INTERVAL_SELECTION);
        table.setRowHeight(22);
        table.setFillsViewportHeight(true);
        table.setShowGrid(true);
        table.setAutoResizeMode(JTable.AUTO_RESIZE_LAST_COLUMN);
        TableColumnModel cm = table.getColumnModel();
        cm.getColumn(0).setMinWidth(34);
        cm.getColumn(0).setMaxWidth(50);
        cm.getColumn(0).setPreferredWidth(42);
        cm.getColumn(1).setPreferredWidth(118);
        cm.getColumn(2).setPreferredWidth(360);
    }

    private static void renumberFileTable(DefaultTableModel model) {
        for (int i = 0; i < model.getRowCount(); i++) {
            model.setValueAt(String.format("%03d", i + 1), i, 0);
        }
    }

    private void clearAllDiffColumns() {
        for (int i = 0; i < beforeTableModel.getRowCount(); i++) {
            beforeTableModel.setValueAt(DIFF_PLACEHOLDER, i, 1);
        }
        for (int i = 0; i < afterTableModel.getRowCount(); i++) {
            afterTableModel.setValueAt(DIFF_PLACEHOLDER, i, 1);
        }
    }

    /**
     * ログエリアにメッセージを追記する。Swing の EDT で実行する。
     *
     * @param message 表示するメッセージ
     */
    private void log(String message) {
        if (logArea == null) return;
        SwingUtilities.invokeLater(() -> {
            logArea.append(message + "\n");
            logArea.setCaretPosition(logArea.getDocument().getLength());
        });
    }

    /**
     * フォルダ入力欄の右側に「フォルダ参照」ボタンを持つ 1 行パネルを作成する。
     *
     * @param labelText ラベルに表示する文字列
     * @param combo     フォルダパス入力用コンボボックス
     * @param isBefore  true のとき移行前、false のとき移行後の入力欄
     * @return 作成したパネル
     */
    private JPanel createFolderRow(String labelText, JComboBox<String> combo, boolean isBefore) {
        JPanel row = new JPanel();
        row.setLayout(new BoxLayout(row, BoxLayout.LINE_AXIS));
        row.add(new JLabel(labelText));
        row.add(Box.createHorizontalStrut(8));
        row.add(combo);
        row.add(Box.createHorizontalStrut(8));
        JButton browseBtn = new JButton("フォルダ参照");
        browseBtn.addActionListener(e -> openSelectedFolderInExplorer(isBefore));
        row.add(browseBtn);
        return row;
    }

    /**
     * リスト見出しの右側に置くタイトル入力欄パネルを作成する。
     *
     * @param titleField タイトル入力欄
     * @return 作成したパネル
     */
    private static JPanel createTitleInputPanel(String sectionLabel, JTextField titleField) {
        JPanel panel = new JPanel(new BorderLayout(6, 2));
        panel.add(new JLabel(sectionLabel), BorderLayout.WEST);
        JPanel right = new JPanel(new FlowLayout(FlowLayout.RIGHT, 4, 0));
        right.add(new JLabel("タイトル:"));
        right.add(titleField);
        panel.add(right, BorderLayout.EAST);
        panel.setBorder(new EmptyBorder(0, 0, 2, 0));
        return panel;
    }
    /**
     * 移行前・移行後の両フォルダから画像ファイルを抽出し、それぞれのリストに反映する。
     */
    private void extractImageFilesFromBoth() {
        extractImageFiles(true);
        extractImageFiles(false);
    }

    /**
     * リストで選択された項目を削除する。複数選択可。インデックスがずれないよう後ろから削除する。
     *
     * @param isBefore true のとき移行前リスト、false のとき移行後リストを対象にする
     */
    private void removeSelectedFromList(boolean isBefore) {
        JTable table = isBefore ? beforeTable : afterTable;
        DefaultTableModel model = isBefore ? beforeTableModel : afterTableModel;
        List<Path> paths = isBefore ? beforeFilePaths : afterFilePaths;
        int[] rows = table.getSelectedRows();
        if (rows.length == 0) {
            log((isBefore ? "画像ファイルA" : "画像ファイルB") + ": 削除する項目を選択してください。");
            return;
        }
        Arrays.sort(rows);
        for (int k = rows.length - 1; k >= 0; k--) {
            int idx = rows[k];
            model.removeRow(idx);
            paths.remove(idx);
        }
        renumberFileTable(model);
        clearAllDiffColumns();
        log((isBefore ? "画像ファイルA" : "画像ファイルB") + ": " + rows.length + " 件を削除しました。");
    }

    /**
     * 指定リストに空白ページ（ラベルのみの項目）を 1 件追加する。
     * 選択行がある場合はその前または後ろに、なければ先頭または末尾に挿入する。
     *
     * @param isBefore     true のとき移行前リスト、false のとき移行後リストを対象にする
     * @param insertBefore true のとき選択行の前に、false のとき選択行の後ろに挿入する
     */
    private void addBlankPageToList(boolean isBefore, boolean insertBefore) {
        JTable table = isBefore ? beforeTable : afterTable;
        DefaultTableModel model = isBefore ? beforeTableModel : afterTableModel;
        List<Path> paths = isBefore ? beforeFilePaths : afterFilePaths;
        int sel = table.getSelectedRow();
        int insertIndex = insertBefore
                ? (sel >= 0 ? sel : 0)
                : (sel >= 0 ? sel + 1 : model.getRowCount());
        model.insertRow(insertIndex, new Object[]{"", DIFF_PLACEHOLDER, BLANK_PAGE_LABEL});
        paths.add(insertIndex, null);
        renumberFileTable(model);
        clearAllDiffColumns();
        log((isBefore ? "画像ファイルA" : "画像ファイルB") + "に空白ページを追加しました。");
    }

    /**
     * 選択した1行を1段上または下へ移動する（隣接行と入れ替え）。複数行選択時はログのみ。
     *
     * @param isBefore true のとき画像ファイルA、false のとき画像ファイルB
     * @param delta    -1 で上へ、+1 で下へ
     */
    private void moveSelectedRowInFileGrid(boolean isBefore, int delta) {
        JTable table = isBefore ? beforeTable : afterTable;
        DefaultTableModel model = isBefore ? beforeTableModel : afterTableModel;
        List<Path> paths = isBefore ? beforeFilePaths : afterFilePaths;
        int[] rows = table.getSelectedRows();
        if (rows.length == 0) {
            log((isBefore ? "画像ファイルA" : "画像ファイルB") + ": 移動する行を選択してください。");
            return;
        }
        if (rows.length != 1) {
            log((isBefore ? "画像ファイルA" : "画像ファイルB") + ": 行の移動は1行だけ選択してください。");
            return;
        }
        int r = rows[0];
        int n = model.getRowCount();
        if (n <= 1) {
            return;
        }
        int other = r + delta;
        if (other < 0 || other >= n) {
            return;
        }
        swapFileTableRows(model, paths, r, other);
        renumberFileTable(model);
        clearAllDiffColumns();
        table.setRowSelectionInterval(other, other);
    }

    private static void swapFileTableRows(DefaultTableModel model, List<Path> paths, int i, int j) {
        for (int col = 0; col < model.getColumnCount(); col++) {
            Object tmp = model.getValueAt(i, col);
            model.setValueAt(model.getValueAt(j, col), i, col);
            model.setValueAt(tmp, j, col);
        }
        Path tmpPath = paths.get(i);
        paths.set(i, paths.get(j));
        paths.set(j, tmpPath);
    }

    /**
     * 移行前または移行後のフォルダパスから画像ファイル一覧を取得し、
     * リストモデルとパスリストを更新する。対象フォルダを履歴に保存する。
     *
     * @param isBefore true のとき移行前、false のとき移行後を対象にする
     */
    private void extractImageFiles(boolean isBefore) {
        String pathStr = comboText(isBefore ? beforeFolderCombo : afterFolderCombo);
        if (pathStr != null) pathStr = pathStr.trim();
        if (pathStr == null || pathStr.isEmpty()) {
            log(isBefore ? "画像フォルダAを指定してください。" : "画像フォルダBを指定してください。");
            return;
        }
        Path dir = Paths.get(pathStr);
        if (!Files.isDirectory(dir)) {
            log("フォルダが存在しません: " + pathStr);
            return;
        }
        List<Path> files;
        try {
            files = listImageFiles(dir);
        } catch (IOException e) {
            log("フォルダの読み込みに失敗しました: " + e.getMessage());
            return;
        }
        files.sort(Comparator.naturalOrder());
        DefaultTableModel model = isBefore ? beforeTableModel : afterTableModel;
        List<Path> paths = isBefore ? beforeFilePaths : afterFilePaths;
        model.setRowCount(0);
        paths.clear();
        for (Path p : files) {
            model.addRow(new Object[]{"", DIFF_PLACEHOLDER, dir.relativize(p).toString()});
            paths.add(p);
        }
        renumberFileTable(model);
        clearAllDiffColumns();
        if (isBefore) saveToHistory("before_folders", pathStr);
        else saveToHistory("after_folders", pathStr);
        log((isBefore ? "画像ファイルA" : "画像ファイルB") + ": " + files.size() + " 件の画像を抽出しました。");
    }

    /**
     * 履歴ファイルから行のリストを読み込む。ファイルが存在しないか読み込み失敗時は空リストを返す。
     *
     * @param fileName 履歴の種類（拡張子 .txt を付与して HISTORY_DIR 内のファイル名になる）
     * @return 空行を除いた行のリスト
     */
    private static List<String> loadHistory(String fileName) {
        Path f = HISTORY_DIR.resolve(fileName + ".txt");
        if (!Files.isRegularFile(f)) return new ArrayList<>();
        try {
            return Files.readAllLines(f, StandardCharsets.UTF_8).stream()
                    .filter(line -> line != null && !line.trim().isEmpty())
                    .collect(Collectors.toList());
        } catch (IOException e) {
            return new ArrayList<>();
        }
    }

    /**
     * コンボボックスの現在の表示テキスト（編集可能時はエディタの内容）を取得する。
     *
     * @param cb 対象のコンボボックス
     * @return 表示文字列。null の場合は空文字
     */
    private static String comboText(JComboBox<String> cb) {
        Object o = cb.isEditable() ? cb.getEditor().getItem() : cb.getSelectedItem();
        return o == null ? "" : o.toString();
    }

    /**
     * 指定された入力欄に設定されたフォルダをエクスプローラーで開く。
     *
     * @param isBefore true のとき移行前、false のとき移行後の入力欄を対象にする
     */
    private void openSelectedFolderInExplorer(boolean isBefore) {
        JComboBox<String> combo = isBefore ? beforeFolderCombo : afterFolderCombo;
        String pathStr = comboText(combo);
        if (pathStr != null) pathStr = pathStr.trim();
        if (pathStr == null || pathStr.isEmpty()) {
            log((isBefore ? "画像フォルダA" : "画像フォルダB") + "のパスを指定してください。");
            return;
        }
        Path dir = Paths.get(pathStr);
        if (!Files.isDirectory(dir)) {
            log("フォルダが存在しません: " + pathStr);
            return;
        }
        openFolderInExplorer(dir);
    }

    /**
     * ボタンのサイズを固定する。
     *
     * @param button 対象のボタン
     * @param width  横幅（px）
     * @param height 高さ（px）
     */
    private static void setFixedButtonSize(JButton button, int width, int height) {
        Dimension fixed = new Dimension(Math.max(0, width), Math.max(0, height));
        button.setPreferredSize(fixed);
        button.setMinimumSize(fixed);
    }

    /**
     * 指定した値を履歴ファイルの先頭に追加する。既存の同じ値は除き、最大 HISTORY_MAX 件まで保持する。
     *
     * @param fileName 履歴の種類（拡張子 .txt を付与して HISTORY_DIR 内のファイル名になる）
     * @param value    保存する文字列（null や空白のみの場合は何もしない）
     */
    private static void saveToHistory(String fileName, String value) {
        if (value == null || value.trim().isEmpty()) return;
        try {
            Files.createDirectories(HISTORY_DIR);
            Path f = HISTORY_DIR.resolve(fileName + ".txt");
            List<String> current = new ArrayList<>();
            if (Files.isRegularFile(f)) {
                current = Files.readAllLines(f, StandardCharsets.UTF_8).stream()
                        .filter(line -> line != null && !line.trim().isEmpty())
                        .collect(Collectors.toList());
            }
            String trimmed = value.trim();
            List<String> updated = new ArrayList<>();
            updated.add(trimmed);
            updated.addAll(current.stream()
                    .filter(line -> !line.equals(trimmed))
                    .limit(HISTORY_MAX - 1)
                    .collect(Collectors.toList()));
            Files.write(f, updated, StandardCharsets.UTF_8);
        } catch (IOException ignored) {
        }
    }

    /**
     * リスト全件を対象に A/B 画像を比較し、差分オーバーレイ結果を PDF にまとめて保存する。
     * 中間画像（PNG）は {@code diff_overlay} フォルダに保存し、PDF は親ディレクトリ直下に作成する。
     */
    private void createDiffOverlayPdf() {
        int maxCount = Math.max(beforeFilePaths.size(), afterFilePaths.size());
        if (maxCount == 0) {
            log("差分PDF: ファイルリストが空です。画像フォルダA・画像フォルダBで「ファイル抽出」を実行してください。");
            return;
        }
        Path outDir = resolveDiffOutputDirectory();
        if (outDir == null) {
            log("差分PDF: 出力先を決めるため、画像フォルダAまたはBを指定してください。");
            return;
        }
        String beforePathStr = comboText(beforeFolderCombo);
        String afterPathStr = comboText(afterFolderCombo);
        String baseName;
        if (beforePathStr != null && !beforePathStr.trim().isEmpty()) {
            baseName = Paths.get(beforePathStr.trim()).getFileName().toString();
        } else if (afterPathStr != null && !afterPathStr.trim().isEmpty()) {
            baseName = Paths.get(afterPathStr.trim()).getFileName().toString();
        } else {
            log("差分PDF: ベース名の決定に失敗しました。");
            return;
        }
        Path pdfPath = outDir.getParent().resolve(baseName + "_diff_overlay.pdf");
        String beforeTitle = beforeTitleField.getText() == null ? "" : beforeTitleField.getText().trim();
        String afterTitle = afterTitleField.getText() == null ? "" : afterTitleField.getText().trim();

        Path diffCaptionBeforeBase = null;
        Path diffCaptionAfterBase = null;
        String comboBefore = comboText(beforeFolderCombo);
        if (comboBefore != null && !comboBefore.trim().isEmpty()) {
            diffCaptionBeforeBase = Paths.get(comboBefore.trim());
        }
        String comboAfter = comboText(afterFolderCombo);
        if (comboAfter != null && !comboAfter.trim().isEmpty()) {
            diffCaptionAfterBase = Paths.get(comboAfter.trim());
        }
        final Path finalDiffCaptionBeforeBase = diffCaptionBeforeBase;
        final Path finalDiffCaptionAfterBase = diffCaptionAfterBase;

        ImageDiffOverlay.DiffSettings settings;
        try {
            settings = readDiffSettingsFromUi();
        } catch (IllegalArgumentException ex) {
            log("差分PDF: " + ex.getMessage());
            JOptionPane.showMessageDialog(null, ex.getMessage(),
                    "差分PDF作成", JOptionPane.WARNING_MESSAGE);
            return;
        }

        diffOverlayButton.setEnabled(false);
        log("差分PDFを作成しています...");

        new Thread(() -> {
            try {
                Files.createDirectories(outDir);
                cleanDiffOverlayWorkFiles(outDir);

                List<Path> beforeDiffPaths = new ArrayList<>();
                List<Path> afterDiffPaths = new ArrayList<>();
                List<String> beforeCaptionOverrides = new ArrayList<>();
                List<String> afterCaptionOverrides = new ArrayList<>();
                List<String> pairSubtitles = new ArrayList<>();
                String[] diffByRow = new String[maxCount];
                Arrays.fill(diffByRow, DIFF_PLACEHOLDER);

                int pairedDiffCount = 0;
                int singleSideRowCount = 0;
                int blankPairRowCount = 0;
                int resizedCount = 0;

                for (int i = 0; i < maxCount; i++) {
                    Path pathA = i < beforeFilePaths.size() ? beforeFilePaths.get(i) : null;
                    Path pathB = i < afterFilePaths.size() ? afterFilePaths.get(i) : null;
                    boolean aOk = pathA != null && Files.isRegularFile(pathA);
                    boolean bOk = pathB != null && Files.isRegularFile(pathB);

                    if (aOk && bOk) {
                        Path outA = outDir.resolve(String.format("diff_%03d_A.png", i + 1));
                        Path outB = outDir.resolve(String.format("diff_%03d_B.png", i + 1));
                        ImageDiffOverlay.MarkedPairResult result = ImageDiffOverlay.writeMarkedPair(
                                pathA, pathB, outA, outB, settings);
                        if (result.resizeNote != null) {
                            resizedCount++;
                        }
                        beforeDiffPaths.add(outA);
                        afterDiffPaths.add(outB);
                        beforeCaptionOverrides.add(toRelativeOrFileName(pathA, finalDiffCaptionBeforeBase));
                        afterCaptionOverrides.add(toRelativeOrFileName(pathB, finalDiffCaptionAfterBase));
                        pairSubtitles.add(formatDiffSubtitleForPdf(result));
                        diffByRow[i] = formatDiffCellForPair(result);
                        pairedDiffCount++;
                    } else if (aOk) {
                        beforeDiffPaths.add(pathA);
                        afterDiffPaths.add(null);
                        beforeCaptionOverrides.add(toRelativeOrFileName(pathA, finalDiffCaptionBeforeBase));
                        afterCaptionOverrides.add(null);
                        pairSubtitles.add(null);
                        singleSideRowCount++;
                    } else if (bOk) {
                        beforeDiffPaths.add(null);
                        afterDiffPaths.add(pathB);
                        beforeCaptionOverrides.add(null);
                        afterCaptionOverrides.add(toRelativeOrFileName(pathB, finalDiffCaptionAfterBase));
                        pairSubtitles.add(null);
                        singleSideRowCount++;
                    } else {
                        beforeDiffPaths.add(null);
                        afterDiffPaths.add(null);
                        beforeCaptionOverrides.add(null);
                        afterCaptionOverrides.add(null);
                        pairSubtitles.add(null);
                        blankPairRowCount++;
                    }
                }

                buildPdfFromLists(
                        beforeDiffPaths, afterDiffPaths, pdfPath.toString(),
                        beforeTitle, afterTitle, outDir, outDir, pairSubtitles,
                        beforeCaptionOverrides, afterCaptionOverrides);

                int finalPaired = pairedDiffCount;
                int finalSingleSide = singleSideRowCount;
                int finalBlankPair = blankPairRowCount;
                int finalResizedCount = resizedCount;
                final String[] diffSnapshot = Arrays.copyOf(diffByRow, diffByRow.length);
                final int finalMax = maxCount;
                SwingUtilities.invokeLater(() -> {
                    diffOverlayButton.setEnabled(true);
                    log("差分PDFを作成しました: " + pdfPath);
                    log(String.format(Locale.US,
                            "差分オーバーレイ（A/B 両方）: %d 件 / 片側のみ画像: %d 行 / 両方なし: %d 行 / リサイズ比較: %d 件",
                            finalPaired, finalSingleSide, finalBlankPair, finalResizedCount));
                    for (int i = 0; i < finalMax; i++) {
                        String d = diffSnapshot[i];
                        if (i < beforeTableModel.getRowCount()) {
                            beforeTableModel.setValueAt(d, i, 1);
                        }
                        if (i < afterTableModel.getRowCount()) {
                            afterTableModel.setValueAt(d, i, 1);
                        }
                    }
                    openFolderInExplorer(pdfPath.getParent());
                });
            } catch (Exception e) {
                e.printStackTrace();
                SwingUtilities.invokeLater(() -> {
                    diffOverlayButton.setEnabled(true);
                    log("差分PDF作成エラー: " + e.getMessage());
                    JOptionPane.showMessageDialog(null, e.getMessage(),
                            "差分PDF作成", JOptionPane.ERROR_MESSAGE);
                });
            }
        }).start();
    }

    /**
     * 画面の差分設定入力欄から値を取得し、妥当性チェック後に返す。
     */
    private ImageDiffOverlay.DiffSettings readDiffSettingsFromUi() {
        int threshold;
        float alpha;
        int expandRadius;
        try {
            threshold = Integer.parseInt(diffThresholdField.getText().trim());
        } catch (Exception e) {
            throw new IllegalArgumentException("差分しきい値は整数で入力してください。");
        }
        try {
            alpha = Float.parseFloat(diffAlphaField.getText().trim());
        } catch (Exception e) {
            throw new IllegalArgumentException("赤αは 0.0〜1.0 の数値で入力してください。");
        }
        try {
            expandRadius = Integer.parseInt(diffExpandRadiusField.getText().trim());
        } catch (Exception e) {
            throw new IllegalArgumentException("マスク拡張は 0 以上の整数で入力してください。");
        }
        if (threshold < 0) {
            throw new IllegalArgumentException("差分しきい値は 0 以上にしてください。");
        }
        if (alpha < 0f || alpha > 1f) {
            throw new IllegalArgumentException("赤αは 0.0〜1.0 の範囲にしてください。");
        }
        if (expandRadius < 0) {
            throw new IllegalArgumentException("マスク拡張は 0 以上にしてください。");
        }
        return new ImageDiffOverlay.DiffSettings(
                threshold, alpha, expandRadius,
                diffOverlayCheckA.isSelected(), diffOverlayCheckB.isSelected());
    }

    /**
     * 差分 PNG の保存先ディレクトリ（{@code diff_overlay}）を返す。
     * 画像フォルダAの親が取れる場合はその直下、無ければBの親の直下。
     *
     * @return 保存先フォルダ。決められない場合は null
     */
    private Path resolveDiffOutputDirectory() {
        String beforePathStr = comboText(beforeFolderCombo);
        if (beforePathStr != null && !beforePathStr.trim().isEmpty()) {
            Path beforeFolder = Paths.get(beforePathStr.trim());
            Path parent = beforeFolder.getParent();
            if (parent != null) {
                return parent.resolve("diff_overlay");
            }
        }
        String afterPathStr = comboText(afterFolderCombo);
        if (afterPathStr != null && !afterPathStr.trim().isEmpty()) {
            Path afterFolder = Paths.get(afterPathStr.trim());
            Path parent = afterFolder.getParent();
            if (parent != null) {
                return parent.resolve("diff_overlay");
            }
        }
        return null;
    }

    /**
     * 移行前・移行後のリストを元に PDF を生成する。
     * 出力先は移行前フォルダの親ディレクトリに「移行前フォルダ名.pdf」で保存し、
     * 作成後にエクスプローラーでそのフォルダを開く。処理は別スレッドで実行する。
     */
    private void createPdf() {
        if (beforeFilePaths.isEmpty() && afterFilePaths.isEmpty()) {
            log("ファイルリストが空です。画像フォルダA・画像フォルダBで「ファイル抽出」を実行してください。");
            return;
        }
        int maxCount = Math.max(beforeFilePaths.size(), afterFilePaths.size());
        if (maxCount == 0) {
            log("画像ファイルが1件もありません。");
            return;
        }

        Path outputDir;
        String baseName;
        Path beforeBaseDir = null;
        Path afterBaseDir = null;
        if (!beforeFilePaths.isEmpty()) {
            String beforePathStr = comboText(beforeFolderCombo);
            if (beforePathStr != null) beforePathStr = beforePathStr.trim();
            if (beforePathStr == null || beforePathStr.isEmpty()) {
                log("PDFの出力先を決めるため、画像フォルダAを指定してください。");
                return;
            }
            Path beforeFolderPath = Paths.get(beforePathStr);
            beforeBaseDir = beforeFolderPath;
            outputDir = beforeFolderPath.getParent();
            baseName = beforeFolderPath.getFileName().toString();
        } else {
            String afterPathStr = comboText(afterFolderCombo);
            if (afterPathStr != null) afterPathStr = afterPathStr.trim();
            if (afterPathStr == null || afterPathStr.isEmpty()) {
                log("PDFの出力先を決めるため、画像フォルダBを指定してください。");
                return;
            }
            Path afterFolderPath = Paths.get(afterPathStr);
            afterBaseDir = afterFolderPath;
            outputDir = afterFolderPath.getParent();
            baseName = afterFolderPath.getFileName().toString();
        }
        if (beforeBaseDir == null) {
            String beforePathStr = comboText(beforeFolderCombo);
            if (beforePathStr != null && !beforePathStr.trim().isEmpty()) {
                beforeBaseDir = Paths.get(beforePathStr.trim());
            }
        }
        if (afterBaseDir == null) {
            String afterPathStr = comboText(afterFolderCombo);
            if (afterPathStr != null && !afterPathStr.trim().isEmpty()) {
                afterBaseDir = Paths.get(afterPathStr.trim());
            }
        }
        String outputPath = outputDir.resolve(baseName + ".pdf").toString();
        String beforeTitle = beforeTitleField.getText() == null ? "" : beforeTitleField.getText().trim();
        String afterTitle = afterTitleField.getText() == null ? "" : afterTitleField.getText().trim();
        final Path finalBeforeBaseDir = beforeBaseDir;
        final Path finalAfterBaseDir = afterBaseDir;

        createPdfButton.setEnabled(false);
        log("PDFを作成しています...");

        new Thread(() -> {
            try {
                buildPdfFromLists(beforeFilePaths, afterFilePaths, outputPath, beforeTitle, afterTitle,
                        finalBeforeBaseDir, finalAfterBaseDir, null, null, null);
                SwingUtilities.invokeLater(() -> {
                    createPdfButton.setEnabled(true);
                    log("PDFを作成しました: " + outputPath);
                    openFolderInExplorer(Paths.get(outputPath).getParent());
                });
            } catch (Exception e) {
                e.printStackTrace();
                SwingUtilities.invokeLater(() -> {
                    createPdfButton.setEnabled(true);
                    log("エラー: " + e.getMessage());
                    JOptionPane.showMessageDialog(null, "エラー: " + e.getMessage(),
                            "エラー", JOptionPane.ERROR_MESSAGE);
                });
            }
        }).start();
    }

    /**
     * デスクトップの機能で指定フォルダをエクスプローラー（または OS のファイルマネージャ）で開く。
     *
     * @param folder 開くフォルダのパス
     */
    private static void openFolderInExplorer(Path folder) {
        try {
            if (Desktop.isDesktopSupported()) {
                Desktop.getDesktop().open(folder.toFile());
            }
        } catch (Exception ignored) {
        }
    }

    /**
     * 移行前・移行後のファイルパスリストから A4 PDF を生成する。
     * 両方に要素がある場合は奇数ページに移行前・偶数ページに移行後を交互に配置する。
     * 片方だけ0件の場合は、もう片方のリストのみを連続して出力する。
     *
     * @param beforeFiles 移行前の画像パス（null 要素は空白ページ）
     * @param afterFiles  移行後の画像パス（null 要素は空白ページ）
     * @param outputPath  出力 PDF のフルパス
     * @param pairSubtitles 交互レイアウト時はインデックスごとに A/B 両ページへ同じ補足行を付ける。{@code null} 可
     * @param beforeCaptionPathOverrides 各行の画像上キャプション（ファイルパス部分）を上書きする。{@code null} 可
     * @param afterCaptionPathOverrides 同上（移行後側）。{@code null} 可
     * @throws Exception PDF の生成・書き込みに失敗した場合
     */
    private void buildPdfFromLists(
            List<Path> beforeFiles, List<Path> afterFiles, String outputPath,
            String beforeTitle, String afterTitle,
            Path beforeBaseDir, Path afterBaseDir,
            List<String> pairSubtitles,
            List<String> beforeCaptionPathOverrides,
            List<String> afterCaptionPathOverrides)
            throws Exception {
        int beforeSize = beforeFiles.size();
        int afterSize = afterFiles.size();
        if (beforeSize == 0 && afterSize == 0) {
            throw new IllegalStateException("画像ファイルが1件もありません。");
        }

        Document doc = new Document(PageSize.A4, 18, 18, 18, 18);
        PdfWriter.getInstance(doc, new FileOutputStream(outputPath));
        doc.open();

        float pageWidth = doc.getPageSize().getWidth() - doc.leftMargin() - doc.rightMargin();
        float pageHeight = doc.getPageSize().getHeight() - doc.topMargin() - doc.bottomMargin();
        double pageW = pageWidth;
        double pageH = pageHeight;

        if (afterSize == 0) {
            // 移行後のみ0件 → 移行前だけを連続で出力
            for (int i = 0; i < beforeSize; i++) {
                if (i > 0) doc.newPage();
                Path path = beforeFiles.get(i);
                if (path != null) {
                    addImagePage(doc, path, pageW, pageH, beforeTitle, beforeBaseDir,
                            pairSubtitleAt(pairSubtitles, i),
                            captionOverrideAt(beforeCaptionPathOverrides, i));
                } else {
                    addBlankPage(doc);
                }
            }
        } else if (beforeSize == 0) {
            // 移行前のみ0件 → 移行後だけを連続で出力
            for (int i = 0; i < afterSize; i++) {
                if (i > 0) doc.newPage();
                Path path = afterFiles.get(i);
                if (path != null) {
                    addImagePage(doc, path, pageW, pageH, afterTitle, afterBaseDir,
                            pairSubtitleAt(pairSubtitles, i),
                            captionOverrideAt(afterCaptionPathOverrides, i));
                } else {
                    addBlankPage(doc);
                }
            }
        } else {
            // 両方に要素あり → 奇数ページ: 移行前、偶数ページ: 移行後で交互に出力
            int maxCount = Math.max(beforeSize, afterSize);
            for (int i = 0; i < maxCount; i++) {
                if (i > 0) doc.newPage();
                String sub = pairSubtitleAt(pairSubtitles, i);
                if (i < beforeSize) {
                    Path path = beforeFiles.get(i);
                    if (path != null) {
                        addImagePage(doc, path, pageW, pageH, beforeTitle, beforeBaseDir, sub,
                                captionOverrideAt(beforeCaptionPathOverrides, i));
                    } else {
                        addBlankPage(doc);
                    }
                } else {
                    addBlankPage(doc);
                }
                doc.newPage();
                if (i < afterSize) {
                    Path path = afterFiles.get(i);
                    if (path != null) {
                        addImagePage(doc, path, pageW, pageH, afterTitle, afterBaseDir, sub,
                                captionOverrideAt(afterCaptionPathOverrides, i));
                    } else {
                        addBlankPage(doc);
                    }
                } else {
                    addBlankPage(doc);
                }
            }
        }

        doc.close();
    }

    /**
     * 指定ディレクトリ配下の画像ファイル（IMAGE_EXTENSIONS に含まれる拡張子）のパスを一覧で返す。
     *
     * @param dir 検索対象のディレクトリ
     * @return 画像ファイルの Path のリスト
     * @throws IOException ディレクトリの読み込みに失敗した場合
     */
    private static List<Path> listImageFiles(Path dir) throws IOException {
        Set<String> lower = new HashSet<>();
        for (String ext : IMAGE_EXTENSIONS) {
            lower.add(ext.toLowerCase());
        }
        return Files.walk(dir)
                .filter(Files::isRegularFile)
                .filter(p -> lower.contains(getExtension(p.getFileName().toString()).toLowerCase()))
                .collect(Collectors.toList());
    }

    /**
     * ファイル名から拡張子（ドット含む）を取得する。
     *
     * @param name ファイル名
     * @return 拡張子（例: ".png"）。ドットがない場合は空文字
     */
    private static String getExtension(String name) {
        int i = name.lastIndexOf('.');
        return i >= 0 ? name.substring(i) : "";
    }

    /**
     * PDF にファイル名・空白ページラベルなどを描画する際に使うフォントを返す。
     * 日本語対応 TTF を優先（リソース → ユーザーディレクトリ → Windows Fonts）し、失敗時は Helvetica。
     */
    private static Font getPdfTextFont() {
        if (pdfTextFont != null) {
            return pdfTextFont;
        }
        try {
            // 1. リソースのフォント（src/main/resources/fonts/ に .ttf を配置）
            String[] resourcePaths = {"fonts/NotoSansJP-Regular.ttf", "fonts/japanese.ttf", "fonts/ipag.ttf", "fonts/ipaexg.ttf"};
            for (String resPath : resourcePaths) {
                try (InputStream in = MigrationScreenshotPdfApp.class.getResourceAsStream("/" + resPath)) {
                    if (in != null) {
                        pdfTextFont = createFontFromStream(in, 10);
                        if (pdfTextFont != null) return pdfTextFont;
                    }
                }
            }
            // 2. ユーザーディレクトリのフォント（.screenShotToPdf/fonts/ に .ttf を配置すると利用可能）
            Path userFontsDir = HISTORY_DIR.resolve("fonts");
            if (Files.isDirectory(userFontsDir)) {
                try (Stream<Path> list = Files.list(userFontsDir)) {
                    Path ttf = list.filter(p -> Files.isRegularFile(p) && p.toString().toLowerCase().endsWith(".ttf"))
                            .findFirst().orElse(null);
                    if (ttf != null) {
                        BaseFont bf = BaseFont.createFont(ttf.toAbsolutePath().toString(), BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
                        pdfTextFont = new Font(bf, 10, Font.NORMAL);
                        return pdfTextFont;
                    }
                } catch (Exception ignored) {
                }
            }
            // 3. Windows の日本語フォント（.ttf のみ。OpenPDF は .ttc を扱えない場合がある）
            String winDir = System.getenv("SystemRoot");
            if (winDir != null) {
                Path fontsDir = Paths.get(winDir, "Fonts");
                String[] winFonts = {"msgothic.ttc", "meiryo.ttc", "yugothm.ttc"};
                for (String name : winFonts) {
                    Path fontPath = fontsDir.resolve(name);
                    if (!Files.isRegularFile(fontPath)) continue;
                    for (String spec : new String[]{fontPath.toAbsolutePath() + ",0", fontPath.toAbsolutePath().toString()}) {
                        try {
                            BaseFont bf = BaseFont.createFont(spec, BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
                            pdfTextFont = new Font(bf, 10, Font.NORMAL);
                            return pdfTextFont;
                        } catch (Exception e) {
                            continue;
                        }
                    }
                }
            }
        } catch (Exception ignored) {
        }
        pdfTextFont = new Font(Font.HELVETICA, 10, Font.NORMAL);
        return pdfTextFont;
    }

    private static Font createFontFromStream(InputStream in, int size) {
        try {
            Path temp = Files.createTempFile("pdffont", ".ttf");
            temp.toFile().deleteOnExit();
            Files.copy(in, temp, StandardCopyOption.REPLACE_EXISTING);
            BaseFont bf = BaseFont.createFont(temp.toAbsolutePath().toString(), BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
            return new Font(bf, size, Font.NORMAL);
        } catch (Exception e) {
            return null;
        }
    }

    /**
     * PDF ドキュメントに空白ページ（ラベルテキストのみ）を 1 ページ追加する。
     *
     * @param doc 対象の PDF ドキュメント
     */
    private static void addBlankPage(Document doc) {
        doc.add(new Paragraph(BLANK_PAGE_LABEL, getPdfTextFont()));
    }

    /**
     * PDF ドキュメントに画像 1 枚を 1 ページとして追加する。
     * ページ内に収まるようスケールし、ファイル名を画像の上に表示する。
     *
     * @param doc       対象の PDF ドキュメント
     * @param imagePath 画像ファイルのパス
     * @param pageWidth ページ内の利用可能幅（pt）
     * @param pageHeight ページ内の利用可能高さ（pt）
     * @param captionPathOverride 非 null のとき、1行目のファイルパス表示をこれにする（画像は {@code imagePath} のまま）
     */
    private void addImagePage(Document doc, Path imagePath, Double pageWidth, Double pageHeight, String title, Path baseDir,
            String subtitle, String captionPathOverride) {
        String displayPath = captionPathOverride != null && !captionPathOverride.isEmpty()
                ? captionPathOverride
                : toRelativeOrFileName(imagePath, baseDir);
        String displayName = formatDisplayName(displayPath, title);
        Font font = getPdfTextFont();
        float imageAreaHeight = pageHeight.floatValue() - 18f;
        if (subtitle != null && !subtitle.trim().isEmpty()) {
            imageAreaHeight -= 14f;
        }

        try {
            Image img = Image.getInstance(imagePath.toAbsolutePath().toString());
            float iw = img.getWidth();
            float ih = img.getHeight();
            if (iw <= 0 || ih <= 0) {
                addBlankPage(doc);
                return;
            }

            float scale = (float) Math.min(pageWidth / iw, imageAreaHeight / ih);
            float w = (float) (iw * scale);
            float h = (float) (ih * scale);
            img.scaleToFit(w, h);
            img.setAlignment(Image.ALIGN_CENTER);

            Paragraph block = new Paragraph();
            block.add(new Chunk(displayName, font));
            block.add(Chunk.NEWLINE);
            if (subtitle != null && !subtitle.trim().isEmpty()) {
                block.add(new Chunk(subtitle.trim(), font));
                block.add(Chunk.NEWLINE);
            }
            block.add(img);
            doc.add(block);
        } catch (Exception e) {
            log("画像読み込みエラー: " + imagePath.toAbsolutePath() + " - " + e.getMessage());
            addBlankPage(doc);
        }
    }

    /**
     * 基準フォルダからの相対パスを返す。相対化できない場合はファイル名を返す。
     */
    private static String toRelativeOrFileName(Path imagePath, Path baseDir) {
        if (imagePath == null) return "";
        try {
            if (baseDir != null) {
                Path normalizedBase = baseDir.toAbsolutePath().normalize();
                Path normalizedImage = imagePath.toAbsolutePath().normalize();
                if (normalizedImage.startsWith(normalizedBase)) {
                    return normalizedBase.relativize(normalizedImage).toString();
                }
            }
        } catch (Exception ignored) {
        }
        Path fileName = imagePath.getFileName();
        return fileName == null ? imagePath.toString() : fileName.toString();
    }

    /**
     * 画像上部に表示する文字列を作る。タイトルが空でない場合は【タイトル】を先頭に付ける。
     *
     * @param fileName ファイル名
     * @param title    タイトル
     * @return 表示文字列
     */
    private static String formatDisplayName(String fileName, String title) {
        String normalizedTitle = title == null ? "" : title.trim();
        if (normalizedTitle.isEmpty()) {
            return fileName;
        }
        return "【" + normalizedTitle + "】 " + fileName;
    }

    private static String pairSubtitleAt(List<String> pairSubtitles, int index) {
        if (pairSubtitles == null || index < 0 || index >= pairSubtitles.size()) {
            return null;
        }
        String s = pairSubtitles.get(index);
        return (s == null || s.isEmpty()) ? null : s;
    }

    private static String captionOverrideAt(List<String> overrides, int index) {
        if (overrides == null || index < 0 || index >= overrides.size()) {
            return null;
        }
        String s = overrides.get(index);
        return (s == null || s.isEmpty()) ? null : s;
    }

    /**
     * 差分PNG出力フォルダ内の、前回実行の diff_*.png / tmp_diff_*.png を削除する。
     */
    private static void cleanDiffOverlayWorkFiles(Path outDir) throws IOException {
        if (!Files.isDirectory(outDir)) {
            return;
        }
        try (Stream<Path> s = Files.list(outDir)) {
            List<Path> toDelete = s.filter(p -> {
                String n = p.getFileName().toString();
                return n.endsWith(".png") && (n.startsWith("diff_") || n.startsWith("tmp_diff_"));
            }).collect(Collectors.toList());
            for (Path p : toDelete) {
                Files.deleteIfExists(p);
            }
        }
    }

    private static String formatDiffSubtitleForPdf(ImageDiffOverlay.MarkedPairResult r) {
        int denom = r.width * r.height;
        double pct = denom > 0 ? 100.0 * r.expandedDiffPixelCount / denom : 0;
        return String.format(Locale.US,
                "差分: %,d px (%.2f%%)",
                r.expandedDiffPixelCount, pct);
    }

    /** 表の「差分」列用（拡張マスク後のピクセル数と画像に対する割合） */
    private static String formatDiffCellForPair(ImageDiffOverlay.MarkedPairResult r) {
        int denom = r.width * r.height;
        double pct = denom > 0 ? 100.0 * r.expandedDiffPixelCount / denom : 0;
        return String.format(Locale.US, "%,d px (%.2f%%)", r.expandedDiffPixelCount, pct);
    }
}

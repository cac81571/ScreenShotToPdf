package com.image2pdf;

import com.formdev.flatlaf.FlatLightLaf;
import com.lowagie.text.Chunk;
import com.lowagie.text.Document;
import com.lowagie.text.Font;
import com.lowagie.text.Image;
import com.lowagie.text.PageSize;
import com.lowagie.text.Paragraph;
import com.lowagie.text.pdf.PdfWriter;

import javax.swing.*;
import javax.swing.border.EmptyBorder;
import javax.swing.border.TitledBorder;
import java.awt.*;
import java.awt.datatransfer.Clipboard;
import java.awt.datatransfer.StringSelection;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Comparator;
import java.util.HashSet;
import java.util.List;
import java.util.Set;
import java.util.stream.Collectors;

/**
 * システムマイグレーションのスクリーンショットを
 * 移行前・移行後で比較できるPDFにまとめるSwingアプリケーション。
 * 奇数ページ＝移行前、偶数ページ＝移行後で交互に配置する。
 */
public class MigrationScreenshotPdfApp {

    /** 空白ページとしてリストに表示するラベル文字列 */
    private static final String BLANK_PAGE_LABEL = "(空白ページ)";
    /** 画像として扱うファイル拡張子の一覧 */
    private static final List<String> IMAGE_EXTENSIONS = Arrays.asList(
            ".png", ".jpg", ".jpeg", ".gif", ".bmp", ".webp");
    /** フォルダ履歴に保持する最大件数 */
    private static final int HISTORY_MAX = 30;
    /** 履歴ファイルを保存するディレクトリ（ユーザーホーム直下の .screenShotToPdf） */
    private static final Path HISTORY_DIR = Paths.get(
            System.getProperty("user.home"), ".screenShotToPdf");
    /** AIプロンプトファイルを格納するディレクトリ */
    private static final Path AI_PROMPT_DIR = HISTORY_DIR.resolve("aiPrompt");
    /** AIプロンプトコンボボックスの固定幅（ピクセル） */
    private static final int AI_PROMPT_COMBO_WIDTH = 300;
    /** 所定フォルダに配置するAIプロンプトのサンプルファイル名（リソース名） */
    private static final List<String> AI_PROMPT_SAMPLE_NAMES = Arrays.asList(
            "マイグレーション比較チェック.txt",
            "サンプル.txt");

    /** 移行前フォルダのパス入力・選択用コンボボックス */
    private JComboBox<String> beforeFolderCombo;
    /** 移行後フォルダのパス入力・選択用コンボボックス */
    private JComboBox<String> afterFolderCombo;
    /** PDF作成を実行するボタン */
    private JButton createPdfButton;
    /** AIプロンプトファイル選択用コンボボックス */
    private JComboBox<String> aiPromptCombo;
    /** ログメッセージを表示するテキストエリア */
    private JTextArea logArea;
    /** 移行前の画像ファイル名リストのモデル */
    private DefaultListModel<String> beforeListModel;
    /** 移行後の画像ファイル名リストのモデル */
    private DefaultListModel<String> afterListModel;
    /** 移行前の画像ファイル一覧を表示するリスト */
    private JList<String> beforeList;
    /** 移行後の画像ファイル一覧を表示するリスト */
    private JList<String> afterList;
    /** 移行前リストに対応するファイルパス（null は空白ページ） */
    private List<Path> beforeFilePaths = new ArrayList<>();
    /** 移行後リストに対応するファイルパス（null は空白ページ） */
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
        List<String> copiedSamples = ensureAiPromptSamples();

        JFrame frame = new JFrame("マイグレーション スクリーンショット PDF 作成ツール");
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        frame.setSize(720, 520);
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
        extractBtn.addActionListener(e -> extractImageFilesFromBoth());

        form.add(createRow("移行前フォルダ:", beforeFolderCombo));
        form.add(Box.createVerticalStrut(8));
        form.add(createRow("移行後フォルダ:", afterFolderCombo));
        form.add(Box.createVerticalStrut(8));
        JPanel extractRow = new JPanel(new BorderLayout());
        extractRow.add(extractBtn, BorderLayout.EAST);
        form.add(extractRow);

        main.add(form, BorderLayout.NORTH);

        beforeListModel = new DefaultListModel<>();
        afterListModel = new DefaultListModel<>();
        beforeList = new JList<>(beforeListModel);
        afterList = new JList<>(afterListModel);
        beforeList.setSelectionMode(ListSelectionModel.MULTIPLE_INTERVAL_SELECTION);
        afterList.setSelectionMode(ListSelectionModel.MULTIPLE_INTERVAL_SELECTION);
        beforeList.setVisibleRowCount(8);
        afterList.setVisibleRowCount(8);

        JPanel listPanel = new JPanel(new GridLayout(1, 2, 10, 0));
        JPanel beforeListPanel = new JPanel(new BorderLayout(2, 4));
        beforeListPanel.setBorder(new TitledBorder(
                new EmptyBorder(2, 4, 2, 4), "移行前 画像ファイル",
                TitledBorder.LEFT, TitledBorder.TOP));
        beforeListPanel.add(new JScrollPane(beforeList), BorderLayout.CENTER);
        JPanel beforeBtnPanel = new JPanel(new FlowLayout(FlowLayout.LEFT, 4, 2));
        JButton beforeDeleteBtn = new JButton("選択削除");
        beforeDeleteBtn.addActionListener(e -> removeSelectedFromList(true));
        JButton beforeBlankBeforeBtn = new JButton("前に空白追加");
        beforeBlankBeforeBtn.addActionListener(e -> addBlankPageToList(true, true));
        JButton beforeBlankAfterBtn = new JButton("後ろに空白追加");
        beforeBlankAfterBtn.addActionListener(e -> addBlankPageToList(true, false));
        beforeBtnPanel.add(beforeDeleteBtn);
        beforeBtnPanel.add(beforeBlankBeforeBtn);
        beforeBtnPanel.add(beforeBlankAfterBtn);
        beforeListPanel.add(beforeBtnPanel, BorderLayout.SOUTH);

        JPanel afterListPanel = new JPanel(new BorderLayout(2, 4));
        afterListPanel.setBorder(new TitledBorder(
                new EmptyBorder(2, 4, 2, 4), "移行後 画像ファイル",
                TitledBorder.LEFT, TitledBorder.TOP));
        afterListPanel.add(new JScrollPane(afterList), BorderLayout.CENTER);
        JPanel afterBtnPanel = new JPanel(new FlowLayout(FlowLayout.LEFT, 4, 2));
        JButton afterDeleteBtn = new JButton("選択削除");
        afterDeleteBtn.addActionListener(e -> removeSelectedFromList(false));
        JButton afterBlankBeforeBtn = new JButton("前に空白追加");
        afterBlankBeforeBtn.addActionListener(e -> addBlankPageToList(false, true));
        JButton afterBlankAfterBtn = new JButton("後ろに空白追加");
        afterBlankAfterBtn.addActionListener(e -> addBlankPageToList(false, false));
        afterBtnPanel.add(afterDeleteBtn);
        afterBtnPanel.add(afterBlankBeforeBtn);
        afterBtnPanel.add(afterBlankAfterBtn);
        afterListPanel.add(afterBtnPanel, BorderLayout.SOUTH);
        listPanel.add(beforeListPanel);
        listPanel.add(afterListPanel);
        main.add(listPanel, BorderLayout.CENTER);

        JPanel bottom = new JPanel(new BorderLayout(0, 6));
        JPanel buttonRow = new JPanel(new BorderLayout(8, 0));
        buttonRow.setAlignmentX(Component.LEFT_ALIGNMENT);
        JPanel aiPromptPanel = new JPanel(new FlowLayout(FlowLayout.LEFT, 8, 0));
        aiPromptCombo = new JComboBox<>(loadAiPromptFiles().toArray(new String[0]));
        aiPromptCombo.setMaximumRowCount(20);
        int comboH = (int) aiPromptCombo.getPreferredSize().getHeight();
        aiPromptCombo.setPreferredSize(new Dimension(AI_PROMPT_COMBO_WIDTH, comboH));
        aiPromptCombo.setMaximumSize(new Dimension(AI_PROMPT_COMBO_WIDTH, comboH));
        if (aiPromptCombo.getItemCount() > 0) {
            aiPromptCombo.setSelectedIndex(0);
        }
        JButton clipboardBtn = new JButton("クリップボード出力");
        clipboardBtn.addActionListener(e -> copySelectedAiPromptToClipboard());
        aiPromptPanel.add(new JLabel("AIプロンプト:"));
        aiPromptPanel.add(aiPromptCombo);
        aiPromptPanel.add(clipboardBtn);
        buttonRow.add(aiPromptPanel, BorderLayout.WEST);
        createPdfButton = new JButton("PDF作成");
        createPdfButton.addActionListener(e -> createPdf());
        buttonRow.add(createPdfButton, BorderLayout.EAST);
        bottom.add(buttonRow, BorderLayout.NORTH);

        logArea = new JTextArea(6, 40);
        logArea.setEditable(false);
        logArea.setLineWrap(true);
        logArea.setWrapStyleWord(true);
        bottom.add(new JScrollPane(logArea), BorderLayout.SOUTH);
        main.add(bottom, BorderLayout.SOUTH);

        for (String name : copiedSamples) {
            log("AIプロンプト サンプルを配置しました: " + name);
        }
        for (int i = 0; i < aiPromptCombo.getItemCount(); i++) {
            String fileName = aiPromptCombo.getItemAt(i);
            String absolutePath = AI_PROMPT_DIR.resolve(fileName).toAbsolutePath().toString();
            log("AIプロンプト リストに追加: " + absolutePath);
        }

        frame.getContentPane().add(main);
        frame.setVisible(true);
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
     * ラベルとコンポーネントを横並びにした 1 行のパネルを作成する。
     *
     * @param labelText ラベルに表示する文字列
     * @param field     右側に配置するコンポーネント
     * @return 作成したパネル
     */
    private static JPanel createRow(String labelText, JComponent field) {
        JPanel row = new JPanel();
        row.setLayout(new BoxLayout(row, BoxLayout.LINE_AXIS));
        row.add(new JLabel(labelText));
        row.add(Box.createHorizontalStrut(8));
        row.add(field);
        return row;
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
        JList<String> list = isBefore ? beforeList : afterList;
        DefaultListModel<String> model = isBefore ? beforeListModel : afterListModel;
        List<Path> paths = isBefore ? beforeFilePaths : afterFilePaths;
        int[] indices = list.getSelectedIndices();
        if (indices.length == 0) {
            log((isBefore ? "移行前" : "移行後") + ": 削除する項目を選択してください。");
            return;
        }
        List<Integer> toRemove = new ArrayList<>();
        for (int idx : indices) {
            toRemove.add(idx);
        }
        toRemove.sort(Comparator.reverseOrder());
        for (int idx : toRemove) {
            model.remove(idx);
            paths.remove(idx);
        }
        log((isBefore ? "移行前" : "移行後") + ": " + toRemove.size() + " 件を削除しました。");
    }

    /**
     * 指定リストに空白ページ（ラベルのみの項目）を 1 件追加する。
     * 選択行がある場合はその前または後ろに、なければ先頭または末尾に挿入する。
     *
     * @param isBefore     true のとき移行前リスト、false のとき移行後リストを対象にする
     * @param insertBefore true のとき選択行の前に、false のとき選択行の後ろに挿入する
     */
    private void addBlankPageToList(boolean isBefore, boolean insertBefore) {
        JList<String> list = isBefore ? beforeList : afterList;
        DefaultListModel<String> model = isBefore ? beforeListModel : afterListModel;
        List<Path> paths = isBefore ? beforeFilePaths : afterFilePaths;
        int insertIndex = insertBefore
                ? (list.getSelectedIndex() >= 0 ? list.getSelectedIndex() : 0)
                : (list.getSelectedIndex() >= 0 ? list.getSelectedIndex() + 1 : model.size());
        model.insertElementAt(BLANK_PAGE_LABEL, insertIndex);
        paths.add(insertIndex, null);
        log((isBefore ? "移行前" : "移行後") + "に空白ページを追加しました。");
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
            log(isBefore ? "移行前フォルダを指定してください。" : "移行後フォルダを指定してください。");
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
        DefaultListModel<String> model = isBefore ? beforeListModel : afterListModel;
        List<Path> paths = isBefore ? beforeFilePaths : afterFilePaths;
        model.clear();
        paths.clear();
        for (Path p : files) {
            model.addElement(p.getFileName().toString());
            paths.add(p);
        }
        if (isBefore) saveToHistory("before_folders", pathStr);
        else saveToHistory("after_folders", pathStr);
        log((isBefore ? "移行前" : "移行後") + ": " + files.size() + " 件の画像を抽出しました。");
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
     * 所定フォルダ（~/.screenShotToPdf/aiPrompt/）を作成し、未配置のサンプルAIプロンプトファイルをリソースからコピーする。既存ファイルは上書きしない。
     *
     * @return コピーしたファイルの絶対パスのリスト（ログ出力用）
     */
    private static List<String> ensureAiPromptSamples() {
        List<String> copied = new ArrayList<>();
        try {
            Files.createDirectories(AI_PROMPT_DIR);
        } catch (IOException e) {
            return copied;
        }
        for (String fileName : AI_PROMPT_SAMPLE_NAMES) {
            Path target = AI_PROMPT_DIR.resolve(fileName);
            if (Files.exists(target)) {
                continue;
            }
            String resourcePath = "/aiPrompt/" + fileName;
            try (InputStream in = MigrationScreenshotPdfApp.class.getResourceAsStream(resourcePath)) {
                if (in != null) {
                    Files.copy(in, target);
                    copied.add(target.toAbsolutePath().toString());
                }
            } catch (IOException ignored) {
            }
        }
        return copied;
    }

    /**
     * AIプロンプト用ディレクトリ（~/.screenShotToPdf/aiPrompt/）配下のファイル名一覧を取得する。
     * ディレクトリが存在しない場合は空リストを返す。
     *
     * @return ファイル名のリスト（ソート済み）
     */
    private static List<String> loadAiPromptFiles() {
        if (!Files.isDirectory(AI_PROMPT_DIR)) {
            return new ArrayList<>();
        }
        try {
            return Files.list(AI_PROMPT_DIR)
                    .filter(Files::isRegularFile)
                    .map(p -> p.getFileName().toString())
                    .sorted()
                    .collect(Collectors.toList());
        } catch (IOException e) {
            return new ArrayList<>();
        }
    }

    /**
     * AIプロンプトコンボで選択されているファイルの内容をクリップボードにコピーする。
     * 未選択またはファイル読み込み失敗時はログにメッセージを出す。
     */
    private void copySelectedAiPromptToClipboard() {
        String fileName = (String) aiPromptCombo.getSelectedItem();
        if (fileName == null || fileName.isEmpty()) {
            log("AIプロンプト: 項目を選択してください。");
            return;
        }
        Path file = AI_PROMPT_DIR.resolve(fileName);
        if (!Files.isRegularFile(file)) {
            log("AIプロンプト: ファイルが見つかりません: " + fileName);
            return;
        }
        try {
            String content = Files.readString(file, StandardCharsets.UTF_8);
            Clipboard clipboard = Toolkit.getDefaultToolkit().getSystemClipboard();
            clipboard.setContents(new StringSelection(content), null);
            log("AIプロンプトをクリップボードにコピーしました: " + fileName);
        } catch (IOException e) {
            log("AIプロンプトの読み込みに失敗しました: " + e.getMessage());
        }
    }

    /**
     * 移行前・移行後のリストを元に PDF を生成する。
     * 出力先は移行前フォルダの親ディレクトリに「移行前フォルダ名.pdf」で保存し、
     * 作成後にエクスプローラーでそのフォルダを開く。処理は別スレッドで実行する。
     */
    private void createPdf() {
        if (beforeFilePaths.isEmpty() && afterFilePaths.isEmpty()) {
            log("ファイルリストが空です。移行前・移行後で「ファイル抽出」を実行してください。");
            return;
        }
        int maxCount = Math.max(beforeFilePaths.size(), afterFilePaths.size());
        if (maxCount == 0) {
            log("画像ファイルが1件もありません。");
            return;
        }

        String beforePathStr = comboText(beforeFolderCombo);
        if (beforePathStr != null) beforePathStr = beforePathStr.trim();
        if (beforePathStr == null || beforePathStr.isEmpty()) {
            log("PDFの出力先を決めるため、移行前フォルダを指定してください。");
            return;
        }
        Path beforeFolderPath = Paths.get(beforePathStr);
        Path outputDir = beforeFolderPath.getParent();
        String baseName = beforeFolderPath.getFileName().toString();
        String outputPath = outputDir.resolve(baseName + ".pdf").toString();

        createPdfButton.setEnabled(false);
        log("PDFを作成しています...");

        new Thread(() -> {
            try {
                buildPdfFromLists(beforeFilePaths, afterFilePaths, outputPath);
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
     * 移行前・移行後のファイルパスリストから、奇数ページに移行前・偶数ページに移行後を配置した
     * A4 PDF を生成する。対応する要素がない側は空白ページで埋める。
     *
     * @param beforeFiles 移行前の画像パス（null 要素は空白ページ）
     * @param afterFiles  移行後の画像パス（null 要素は空白ページ）
     * @param outputPath  出力 PDF のフルパス
     * @throws Exception PDF の生成・書き込みに失敗した場合
     */
    private void buildPdfFromLists(List<Path> beforeFiles, List<Path> afterFiles, String outputPath)
            throws Exception {
        int maxCount = Math.max(beforeFiles.size(), afterFiles.size());
        if (maxCount == 0) {
            throw new IllegalStateException("画像ファイルが1件もありません。");
        }

        Document doc = new Document(PageSize.A4, 18, 18, 18, 18);
        PdfWriter writer = PdfWriter.getInstance(doc, new FileOutputStream(outputPath));
        doc.open();

        float pageWidth = doc.getPageSize().getWidth() - doc.leftMargin() - doc.rightMargin();
        float pageHeight = doc.getPageSize().getHeight() - doc.topMargin() - doc.bottomMargin();

        for (int i = 0; i < maxCount; i++) {
            if (i > 0) doc.newPage();
            // 奇数ページ: 移行前
            if (i < beforeFiles.size()) {
                Path path = beforeFiles.get(i);
                if (path != null) {
                    addImagePage(doc, path.toAbsolutePath().toString(), (double) pageWidth, (double) pageHeight);
                } else {
                    addBlankPage(doc);
                }
            } else {
                addBlankPage(doc);
            }
            doc.newPage();
            // 偶数ページ: 移行後
            if (i < afterFiles.size()) {
                Path path = afterFiles.get(i);
                if (path != null) {
                    addImagePage(doc, path.toAbsolutePath().toString(), (double) pageWidth, (double) pageHeight);
                } else {
                    addBlankPage(doc);
                }
            } else {
                addBlankPage(doc);
            }
        }

        doc.close();
    }

    /**
     * 指定ディレクトリ内の画像ファイル（IMAGE_EXTENSIONS に含まれる拡張子）のパスを一覧で返す。
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
        return Files.list(dir)
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
     * PDF ドキュメントに空白ページ（ラベルテキストのみ）を 1 ページ追加する。
     *
     * @param doc 対象の PDF ドキュメント
     */
    private static void addBlankPage(Document doc) {
        Font font = new Font(Font.HELVETICA, 10, Font.NORMAL);
        doc.add(new Paragraph(BLANK_PAGE_LABEL, font));
    }

    /**
     * PDF ドキュメントに画像 1 枚を 1 ページとして追加する。
     * ページ内に収まるようスケールし、ファイル名を画像の上に表示する。
     *
     * @param doc       対象の PDF ドキュメント
     * @param imagePath 画像ファイルのフルパス
     * @param pageWidth ページ内の利用可能幅（pt）
     * @param pageHeight ページ内の利用可能高さ（pt）
     */
    private void addImagePage(Document doc, String imagePath, Double pageWidth, Double pageHeight) {
        String fileName = Paths.get(imagePath).getFileName().toString();
        Font font = new Font(Font.HELVETICA, 10, Font.NORMAL);
        float imageAreaHeight = pageHeight.floatValue() - 18f;

        try {
            Image img = Image.getInstance(imagePath);
            float iw = img.getWidth();
            float ih = img.getHeight();
            if (iw <= 0 || ih <= 0) return;

            float scale = (float) Math.min(pageWidth / iw, imageAreaHeight / ih);
            float w = (float) (iw * scale);
            float h = (float) (ih * scale);
            img.scaleToFit(w, h);
            img.setAlignment(Image.ALIGN_CENTER);

            Paragraph block = new Paragraph();
            block.add(new Chunk(fileName, font));
            block.add(Chunk.NEWLINE);
            block.add(img);
            doc.add(block);
        } catch (Exception e) {
            log("画像読み込みエラー: " + imagePath + " - " + e.getMessage());
        }
    }
}

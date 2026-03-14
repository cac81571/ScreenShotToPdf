package com.image2pdf

import com.lowagie.text.Chunk
import com.lowagie.text.Document
import com.lowagie.text.Font
import com.lowagie.text.Image
import com.lowagie.text.PageSize
import com.lowagie.text.Paragraph
import com.lowagie.text.pdf.PdfWriter

import com.formdev.flatlaf.FlatLightLaf

import javax.swing.*
import javax.swing.border.EmptyBorder
import javax.swing.border.TitledBorder
import java.awt.*
import java.io.FileOutputStream
import java.nio.file.Files
import java.nio.file.Path
import java.nio.file.Paths

/**
 * システムマイグレーションのスクリーンショットを
 * 移行前・移行後で比較できるPDFにまとめるSwingアプリケーション。
 * 奇数ページ＝移行前、偶数ページ＝移行後で交互に配置する。
 */
class MigrationScreenshotPdfApp {

    static void main(String[] args) {
        FlatLightLaf.setup()
        SwingUtilities.invokeLater { new MigrationScreenshotPdfApp().run() }
    }

    private JComboBox<String> beforeFolderCombo
    private JComboBox<String> afterFolderCombo
    private JButton createPdfButton
    private JTextArea logArea
    private DefaultListModel<String> beforeListModel
    private DefaultListModel<String> afterListModel
    private JList<String> beforeList
    private JList<String> afterList
    private java.util.List<Path> beforeFilePaths = []
    private java.util.List<Path> afterFilePaths = []

    private static final String BLANK_PAGE_LABEL = '(空白ページ)'
    private static final java.util.List<String> IMAGE_EXTENSIONS = ['.png', '.jpg', '.jpeg', '.gif', '.bmp', '.webp']
    private static final int HISTORY_MAX = 30
    private static final Path HISTORY_DIR = Paths.get(System.getProperty('user.home'), '.screenShotToPdf')

    void run() {
        def frame = new JFrame('マイグレーション スクリーンショット → PDF')
        frame.defaultCloseOperation = JFrame.EXIT_ON_CLOSE
        frame.setSize(720, 520)
        frame.setLocationRelativeTo(null)

        def main = new JPanel(new BorderLayout(10, 10))
        main.border = new EmptyBorder(15, 15, 15, 15)

        def form = new JPanel()
        form.layout = new BoxLayout(form, BoxLayout.PAGE_AXIS)
        form.alignmentX = Component.LEFT_ALIGNMENT

        beforeFolderCombo = new JComboBox<>(loadHistory('before_folders').toArray(new String[0]))
        afterFolderCombo = new JComboBox<>(loadHistory('after_folders').toArray(new String[0]))
        [beforeFolderCombo, afterFolderCombo].each { JComboBox<String> cb ->
            cb.editable = true
            int h = (int) cb.preferredSize.height
            cb.maximumSize = new Dimension(Integer.MAX_VALUE, h)
        }

        def extractBtn = new JButton('ファイル抽出')
        extractBtn.addActionListener { extractImageFilesFromBoth() }

        form.add(createRow('移行前フォルダ:', beforeFolderCombo))
        form.add(Box.createVerticalStrut(8))
        form.add(createRow('移行後フォルダ:', afterFolderCombo))
        form.add(Box.createVerticalStrut(8))
        def extractRow = new JPanel(new BorderLayout())
        extractRow.add(extractBtn, BorderLayout.EAST)
        form.add(extractRow)

        main.add(form, BorderLayout.NORTH)

        beforeListModel = new DefaultListModel<>()
        afterListModel = new DefaultListModel<>()
        beforeList = new JList<>(beforeListModel)
        afterList = new JList<>(afterListModel)
        beforeList.selectionMode = ListSelectionModel.MULTIPLE_INTERVAL_SELECTION
        afterList.selectionMode = ListSelectionModel.MULTIPLE_INTERVAL_SELECTION
        beforeList.visibleRowCount = 8
        afterList.visibleRowCount = 8
        def listPanel = new JPanel(new GridLayout(1, 2, 10, 0))
        def beforeListPanel = new JPanel(new BorderLayout(2, 4))
        beforeListPanel.border = new TitledBorder(new EmptyBorder(2, 4, 2, 4), '移行前 画像ファイル', TitledBorder.LEFT, TitledBorder.TOP)
        beforeListPanel.add(new JScrollPane(beforeList), BorderLayout.CENTER)
        def beforeBtnPanel = new JPanel(new FlowLayout(FlowLayout.LEFT, 4, 2))
        def beforeDeleteBtn = new JButton('選択削除')
        beforeDeleteBtn.addActionListener { removeSelectedFromList(true) }
        def beforeBlankBeforeBtn = new JButton('前に空白追加')
        beforeBlankBeforeBtn.addActionListener { addBlankPageToList(true, true) }
        def beforeBlankAfterBtn = new JButton('後ろに空白追加')
        beforeBlankAfterBtn.addActionListener { addBlankPageToList(true, false) }
        beforeBtnPanel.add(beforeDeleteBtn)
        beforeBtnPanel.add(beforeBlankBeforeBtn)
        beforeBtnPanel.add(beforeBlankAfterBtn)
        beforeListPanel.add(beforeBtnPanel, BorderLayout.SOUTH)
        def afterListPanel = new JPanel(new BorderLayout(2, 4))
        afterListPanel.border = new TitledBorder(new EmptyBorder(2, 4, 2, 4), '移行後 画像ファイル', TitledBorder.LEFT, TitledBorder.TOP)
        afterListPanel.add(new JScrollPane(afterList), BorderLayout.CENTER)
        def afterBtnPanel = new JPanel(new FlowLayout(FlowLayout.LEFT, 4, 2))
        def afterDeleteBtn = new JButton('選択削除')
        afterDeleteBtn.addActionListener { removeSelectedFromList(false) }
        def afterBlankBeforeBtn = new JButton('前に空白追加')
        afterBlankBeforeBtn.addActionListener { addBlankPageToList(false, true) }
        def afterBlankAfterBtn = new JButton('後ろに空白追加')
        afterBlankAfterBtn.addActionListener { addBlankPageToList(false, false) }
        afterBtnPanel.add(afterDeleteBtn)
        afterBtnPanel.add(afterBlankBeforeBtn)
        afterBtnPanel.add(afterBlankAfterBtn)
        afterListPanel.add(afterBtnPanel, BorderLayout.SOUTH)
        listPanel.add(beforeListPanel)
        listPanel.add(afterListPanel)
        main.add(listPanel, BorderLayout.CENTER)

        def bottom = new JPanel(new BorderLayout(0, 6))
        def buttonRow = new JPanel(new BorderLayout())
        buttonRow.alignmentX = Component.LEFT_ALIGNMENT
        createPdfButton = new JButton('PDF作成')
        createPdfButton.addActionListener { createPdf() }
        buttonRow.add(createPdfButton, BorderLayout.EAST)
        bottom.add(buttonRow, BorderLayout.NORTH)

        logArea = new JTextArea(6, 40)
        logArea.editable = false
        logArea.lineWrap = true
        logArea.wrapStyleWord = true
        bottom.add(new JScrollPane(logArea), BorderLayout.SOUTH)
        main.add(bottom, BorderLayout.SOUTH)

        frame.contentPane.add(main)
        frame.visible = true
    }

    private void log(String message) {
        if (logArea == null) return
        SwingUtilities.invokeLater {
            logArea.append(message + '\n')
            logArea.caretPosition = logArea.document.length
        }
    }

    private static JPanel createRow(String labelText, JComponent field) {
        def row = new JPanel()
        row.layout = new BoxLayout(row, BoxLayout.LINE_AXIS)
        row.add(new JLabel(labelText))
        row.add(Box.createHorizontalStrut(8))
        row.add(field)
        row
    }

    private void extractImageFilesFromBoth() {
        extractImageFiles(true)
        extractImageFiles(false)
    }

    private void removeSelectedFromList(boolean isBefore) {
        def list = isBefore ? beforeList : afterList
        def model = isBefore ? beforeListModel : afterListModel
        def paths = isBefore ? beforeFilePaths : afterFilePaths
        def indices = list.selectedIndices
        if (indices.length == 0) {
            log((isBefore ? '移行前' : '移行後') + ': 削除する項目を選択してください。')
            return
        }
        def toRemove = indices.toList().sort().reverse()
        toRemove.each { int idx ->
            model.remove(idx)
            paths.remove(idx)
        }
        log((isBefore ? '移行前' : '移行後') + ": ${toRemove.size()} 件を削除しました。")
    }

    private void addBlankPageToList(boolean isBefore, boolean insertBefore) {
        def list = isBefore ? beforeList : afterList
        def model = isBefore ? beforeListModel : afterListModel
        def paths = isBefore ? beforeFilePaths : afterFilePaths
        int insertIndex = insertBefore
            ? (list.selectedIndex >= 0 ? list.selectedIndex : 0)
            : (list.selectedIndex >= 0 ? list.selectedIndex + 1 : model.size())
        model.insertElementAt(BLANK_PAGE_LABEL, insertIndex)
        paths.add(insertIndex, null)
        log((isBefore ? '移行前' : '移行後') + 'に空白ページを追加しました。')
    }

    private void extractImageFiles(boolean isBefore) {
        def pathStr = isBefore ? comboText(beforeFolderCombo)?.trim() : comboText(afterFolderCombo)?.trim()
        if (!pathStr) {
            log(isBefore ? '移行前フォルダを指定してください。' : '移行後フォルダを指定してください。')
            return
        }
        def dir = Paths.get(pathStr)
        if (!Files.isDirectory(dir)) {
            log("フォルダが存在しません: $pathStr")
            return
        }
        def files = listImageFiles(dir).sort()
        def model = isBefore ? beforeListModel : afterListModel
        def paths = isBefore ? beforeFilePaths : afterFilePaths
        model.clear()
        paths.clear()
        files.each { Path p ->
            model.addElement(p.fileName.toString())
            paths.add(p)
        }
        if (isBefore) saveToHistory('before_folders', pathStr)
        else saveToHistory('after_folders', pathStr)
        log((isBefore ? '移行前' : '移行後') + ": ${files.size()} 件の画像を抽出しました。")
    }

    private static java.util.List<String> loadHistory(String fileName) {
        def f = HISTORY_DIR.resolve("${fileName}.txt").toFile()
        if (!f.file) return []
        try {
            return f.readLines('UTF-8').findAll { it?.trim() }
        } catch (Exception e) {
            return []
        }
    }

    private static String comboText(JComboBox<String> cb) {
        def o = cb.editable ? cb.editor.item : cb.selectedItem
        return o == null ? '' : o.toString()
    }

    private static void saveToHistory(String fileName, String value) {
        if (!value?.trim()) return
        try {
            Files.createDirectories(HISTORY_DIR)
            def f = HISTORY_DIR.resolve("${fileName}.txt").toFile()
            def current = f.file ? f.readLines('UTF-8').findAll { it?.trim() } : []
            def updated = [value.trim()] + current.findAll { it != value.trim() }.take(HISTORY_MAX - 1)
            f.withWriter('UTF-8') { w -> updated.each { w.writeLine(it) } }
        } catch (Exception ignored) {}
    }

    private void createPdf() {
        if (beforeFilePaths.isEmpty() && afterFilePaths.isEmpty()) {
            log('ファイルリストが空です。移行前・移行後で「ファイル抽出」を実行してください。')
            return
        }
        def maxCount = Math.max(beforeFilePaths.size(), afterFilePaths.size())
        if (maxCount == 0) {
            log('画像ファイルが1件もありません。')
            return
        }

        def beforePathStr = comboText(beforeFolderCombo)?.trim()
        if (!beforePathStr) {
            log('PDFの出力先を決めるため、移行前フォルダを指定してください。')
            return
        }
        def beforeFolderPath = Paths.get(beforePathStr)
        def outputDir = beforeFolderPath.parent
        def baseName = beforeFolderPath.fileName.toString()
        def outputPath = outputDir.resolve(baseName + '.pdf').toString()

        createPdfButton.enabled = false
        log('PDFを作成しています...')

        Thread.start {
            try {
                buildPdfFromLists(beforeFilePaths, afterFilePaths, outputPath)
                SwingUtilities.invokeLater {
                    createPdfButton.enabled = true
                    log("PDFを作成しました: $outputPath")
                    openFolderInExplorer(Paths.get(outputPath).parent)
                }
            } catch (Exception e) {
                e.printStackTrace()
                SwingUtilities.invokeLater {
                    createPdfButton.enabled = true
                    log("エラー: ${e.message}")
                    JOptionPane.showMessageDialog(null, "エラー: ${e.message}", 'エラー', JOptionPane.ERROR_MESSAGE)
                }
            }
        }
    }

    private static void openFolderInExplorer(Path folder) {
        try {
            if (Desktop.isDesktopSupported()) {
                Desktop.desktop.open(folder.toFile())
            }
        } catch (Exception ignored) {}
    }

    private void buildPdfFromLists(java.util.List<Path> beforeFiles, java.util.List<Path> afterFiles, String outputPath) {
        def maxCount = Math.max(beforeFiles.size(), afterFiles.size())
        if (maxCount == 0) {
            throw new IllegalStateException('画像ファイルが1件もありません。')
        }

        def doc = new Document(PageSize.A4, 18, 18, 18, 18)
        def writer = PdfWriter.getInstance(doc, new FileOutputStream(outputPath))
        doc.open()

        def pageWidth = doc.pageSize.width - doc.leftMargin() - doc.rightMargin()
        def pageHeight = doc.pageSize.height - doc.topMargin() - doc.bottomMargin()

        for (int i = 0; i < maxCount; i++) {
            if (i > 0) doc.newPage()
            // 奇数ページ: 移行前
            if (i < beforeFiles.size()) {
                def path = beforeFiles[i]
                if (path != null) {
                    addImagePage(doc, path.toAbsolutePath().toString(), pageWidth, pageHeight)
                } else {
                    addBlankPage(doc)
                }
            } else {
                addBlankPage(doc)
            }
            doc.newPage()
            // 偶数ページ: 移行後
            if (i < afterFiles.size()) {
                def path = afterFiles[i]
                if (path != null) {
                    addImagePage(doc, path.toAbsolutePath().toString(), pageWidth, pageHeight)
                } else {
                    addBlankPage(doc)
                }
            } else {
                addBlankPage(doc)
            }
        }

        doc.close()
    }

    private static java.util.List<Path> listImageFiles(Path dir) {
        def lower = IMAGE_EXTENSIONS as Set
        return Files.list(dir)
            .filter { Files.isRegularFile(it) }
            .filter { lower.contains(getExtension(it.fileName.toString()).toLowerCase()) }
            .collect { it }
    }

    private static String getExtension(String name) {
        def i = name.lastIndexOf('.')
        return i >= 0 ? name.substring(i) : ''
    }

    private static void addBlankPage(Document doc) {
        def font = new Font(Font.HELVETICA, 10, Font.NORMAL)
        doc.add(new Paragraph(BLANK_PAGE_LABEL, font))
    }

    private void addImagePage(Document doc, String imagePath, Double pageWidth, Double pageHeight) {
        def fileName = Paths.get(imagePath).fileName.toString()
        def font = new Font(Font.HELVETICA, 10, Font.NORMAL)
        // ファイル名用の高さを確保し、画像は残りに収める
        def imageAreaHeight = pageHeight - 18f

        def img = Image.getInstance(imagePath)
        float iw = img.width
        float ih = img.height
        if (iw <= 0 || ih <= 0) return

        float scale = (float) Math.min(pageWidth / iw, imageAreaHeight / ih)
        float w = (float) (iw * scale)
        float h = (float) (ih * scale)
        img.scaleToFit(w, h)
        img.setAlignment(Image.ALIGN_CENTER)

        // ファイル名と画像を1つの Paragraph にまとめて同一ページに収める
        def block = new Paragraph()
        block.add(new Chunk(fileName, font))
        block.add(Chunk.NEWLINE)
        block.add(img)
        doc.add(block)
    }
}

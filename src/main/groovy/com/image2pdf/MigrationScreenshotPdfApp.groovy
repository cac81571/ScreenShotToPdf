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

    private static final java.util.List<String> IMAGE_EXTENSIONS = ['.png', '.jpg', '.jpeg', '.gif', '.bmp', '.webp']
    private static final int HISTORY_MAX = 30
    private static final Path HISTORY_DIR = Paths.get(System.getProperty('user.home'), '.screenShotToPdf')

    void run() {
        def frame = new JFrame('マイグレーション スクリーンショット → PDF')
        frame.defaultCloseOperation = JFrame.EXIT_ON_CLOSE
        frame.setSize(640, 380)
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

        form.add(createRow('移行前フォルダ:', beforeFolderCombo))
        form.add(Box.createVerticalStrut(8))
        form.add(createRow('移行後フォルダ:', afterFolderCombo))

        main.add(form, BorderLayout.NORTH)

        def bottom = new JPanel(new BorderLayout(0, 6))
        def buttonRow = new JPanel(new BorderLayout())
        buttonRow.alignmentX = Component.LEFT_ALIGNMENT
        createPdfButton = new JButton('PDF作成')
        createPdfButton.addActionListener { createPdf() }
        buttonRow.add(createPdfButton, BorderLayout.EAST)
        bottom.add(buttonRow, BorderLayout.NORTH)

        logArea = new JTextArea(8, 40)
        logArea.editable = false
        logArea.lineWrap = true
        logArea.wrapStyleWord = true
        bottom.add(new JScrollPane(logArea), BorderLayout.CENTER)
        main.add(bottom, BorderLayout.CENTER)

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
        def beforePath = comboText(beforeFolderCombo)?.trim()
        def afterPath = comboText(afterFolderCombo)?.trim()

        if (!beforePath) {
            log('移行前フォルダを指定してください。')
            return
        }
        if (!afterPath) {
            log('移行後フォルダを指定してください。')
            return
        }

        def beforeDir = Paths.get(beforePath)
        def afterDir = Paths.get(afterPath)

        if (!Files.isDirectory(beforeDir)) {
            log("移行前フォルダが存在しません: $beforePath")
            return
        }
        if (!Files.isDirectory(afterDir)) {
            log("移行後フォルダが存在しません: $afterPath")
            return
        }

        // 移行前フォルダの親フォルダに「移行前フォルダ名.pdf」で出力
        def outputPath = beforeDir.parent.resolve(beforeDir.fileName.toString() + '.pdf').toString()

        createPdfButton.enabled = false
        log('PDFを作成しています...')

        Thread.start {
            try {
                buildPdf(beforeDir, afterDir, outputPath)
                saveToHistory('before_folders', beforePath)
                saveToHistory('after_folders', afterPath)
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

    private void buildPdf(Path beforeDir, Path afterDir, String outputPath) {
        def beforeFiles = listImageFiles(beforeDir).sort()
        def afterFiles = listImageFiles(afterDir).sort()

        def maxCount = Math.max(beforeFiles.size(), afterFiles.size())
        if (maxCount == 0) {
            throw new IllegalStateException('いずれのフォルダにも画像ファイルが見つかりません。')
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
                addImagePage(doc, beforeFiles[i].toAbsolutePath().toString(), pageWidth, pageHeight)
            } else {
                doc.newPage()
            }
            doc.newPage()
            // 偶数ページ: 移行後
            if (i < afterFiles.size()) {
                addImagePage(doc, afterFiles[i].toAbsolutePath().toString(), pageWidth, pageHeight)
            } else {
                doc.newPage()
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

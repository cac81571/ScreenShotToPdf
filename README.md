# Migration screenshot → PDF

A tool that combines system migration screenshots into a single PDF for easy side-by-side comparison between **image folder A** (before migration) and **image folder B** (after migration). You can also produce a separate PDF that **visualizes pixel differences with a semi-transparent red mask**.

- **Comparison PDF (Create PDF)**
  - **Odd pages**: images from folder A (sorted by filename)
  - **Even pages**: images from folder B (sorted by filename)
- **Diff PDF (Create diff PDF)**
  - For **every row** in the list, compares A/B images at the same index and outputs images with the diff mask overlaid on **alternating pages** (same order as the comparison PDF)
  - Intermediate PNGs are saved under the `diff_overlay` folder; the PDF is saved in its parent folder as `<folderA_name>_diff_overlay.pdf`
- Supported images: PNG, JPG, JPEG, GIF, BMP, WebP

## Tech stack

- **Java 11** + **Swing** (GUI)
- **Maven** (build)
- **OpenPDF** (PDF generation)
- **FlatLaf** (Look and Feel)

## Requirements

- Java 11 or later
- Maven 3.6 or later (for building)

## Build and run

```bash
# Build (produces a runnable JAR with dependencies under target/)
mvn package -DskipTests

# Run (standalone JAR)
java -jar target/migration-screenshot-pdf-1.0.0-SNAPSHOT.jar
```

During development you can also start the app with `mvn exec:java`.

## Usage

1. Launch the app; a Swing window opens.
2. **Image folder A** … Enter or pick the path to the folder with pre-migration screenshots in the combo box.
3. **Image folder B** … Enter or pick the path to the folder with post-migration screenshots in the combo box.
4. Click **Extract files** to scan both folders for images (extensions above) and show them in each list.
5. In the lists you can:
   - **Remove selection** … Remove the selected rows.
   - **Insert blank before** / **Insert blank after** … Insert one “blank page” row before or after the selection (in the PDF this becomes a label-only page).
6. Click **Create PDF** to save a PDF named after folder A as `FolderAName.pdf` in folder A’s **parent directory**, and open that folder in Explorer.
7. Click **Create diff PDF** to run diff processing for all list rows and write `FolderAName_diff_overlay.pdf` in the parent directory. PNGs generated during processing are saved under `diff_overlay` next to the parent.

Images at the same index appear on odd (A) and even (B) pages so you can compare them easily. Folder paths are stored as history under `~/.screenShotToPdf/`.

### Diff tuning (bottom of the window)

| Setting | Description |
|---------|-------------|
| **Diff threshold** | Pixels whose sum of R/G/B differences exceeds this value are treated as different (integer, ≥ 0). Higher values ignore more noise. |
| **Red alpha** | Opacity of the red overlay on differing areas (0.0–1.0). |
| **Mask dilation** | Radius in pixels (≥ 0) to grow diff pixels outward. Larger values make the mask look thicker. Default is 10. |
| **Mask on A** / **Mask on B** | Whether to overlay the mask on each output image. Defaults: **A off, B on**. |

If resolutions differ between A and B, **image B is resized to match image A’s width and height** before comparison.

In the diff PDF, rows where only one side has an image (blank rows or missing files) are skipped; that position becomes a blank page in the PDF.

**Japanese filenames and blank pages**: PDF text supports Japanese display. Provide a Japanese font (TTF) using one of the following:

- **Recommended (reliable on Windows)**  
  1. Download a Japanese-capable **.ttf** font (e.g. [IPAex fonts](https://ipafont.ipa.go.jp/)).  
  2. Create `.screenShotToPdf/fonts/` under your user home and place **one** **.ttf** there (e.g. `C:\Users\<username>\.screenShotToPdf\fonts\ipaexg.ttf`).  
  3. Start the app and create a PDF.

- For development, placing a **.ttf** under `src/main/resources/fonts/` bundles it in the JAR.  
- Windows fonts like Meiryo (.ttc) may not work with OpenPDF; in that case use the TTF setup above.

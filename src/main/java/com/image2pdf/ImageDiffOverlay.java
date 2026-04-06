package com.image2pdf;

import javax.imageio.ImageIO;
import java.awt.Graphics2D;
import java.awt.RenderingHints;
import java.awt.image.BufferedImage;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;

/**
 * 2枚の画像を比較し、差分領域への赤半透明マスクや A/B ブレンドなど、設定に応じた画像を書き出す。
 * 解像度が異なる場合は画像Aの幅・高さに合わせて画像Bをリサイズしてから比較する。
 */
public final class ImageDiffOverlay {

    /** RGB 各チャンネル差の合計がこの値を超えたら差分とみなす（JPEG の軽いノイズを多少吸収） */
    private static final int DEFAULT_RGB_SUM_THRESHOLD = 32;
    /** 差分部分に乗せる赤の不透明度（デフォルト） */
    private static final float DEFAULT_RED_ALPHA = 0.45f;
    /** 差分マスクを太らせる半径（px）。1 なら 3x3、2 なら 5x5 に広がる。 */
    private static final int MASK_EXPAND_RADIUS = 10;
    /** ブレンド時の各画像の重み（合計 1.0 で半透明同士の重ね合わせ） */
    private static final float BLEND_WEIGHT_EACH = 0.5f;

    private ImageDiffOverlay() {
    }

    /** {@link #writeMarkedPair(Path, Path, Path, Path, DiffSettings)} の結果（リサイズ案内と差分の数値） */
    public static final class MarkedPairResult {
        /** B を A に合わせてリサイズしたときのログ用メッセージ。不要な場合は {@code null} */
        public final String resizeNote;
        /** しきい値通過の差分ピクセル数（マスク拡張前） */
        public final int rawDiffPixelCount;
        /** 拡張後マスクの差分ピクセル数（表の「差分」列・統計用） */
        public final int expandedDiffPixelCount;
        public final int width;
        public final int height;

        public MarkedPairResult(String resizeNote, int rawDiffPixelCount, int expandedDiffPixelCount, int width, int height) {
            this.resizeNote = resizeNote;
            this.rawDiffPixelCount = rawDiffPixelCount;
            this.expandedDiffPixelCount = expandedDiffPixelCount;
            this.width = width;
            this.height = height;
        }
    }

    /** 差分描画パラメータ */
    public static final class DiffSettings {
        public final int rgbSumThreshold;
        public final float redAlpha;
        public final int maskExpandRadius;
        /** true のとき出力A（移行前）を A と B のブレンドで出力する（{@link #redMaskOnA} が true なら無視） */
        public final boolean blendDisplayOnA;
        /** true のとき出力B（移行後）を A と B のブレンドで出力する（{@link #redMaskOnB} が true なら無視） */
        public final boolean blendDisplayOnB;
        /** true のとき出力A に拡張差分マスクへ赤半透明を重ねる */
        public final boolean redMaskOnA;
        /** true のとき出力B に拡張差分マスクへ赤半透明を重ねる */
        public final boolean redMaskOnB;

        public DiffSettings(int rgbSumThreshold, float redAlpha, int maskExpandRadius,
                boolean blendDisplayOnA, boolean blendDisplayOnB,
                boolean redMaskOnA, boolean redMaskOnB) {
            this.rgbSumThreshold = rgbSumThreshold;
            this.redAlpha = redAlpha;
            this.maskExpandRadius = maskExpandRadius;
            this.blendDisplayOnA = blendDisplayOnA;
            this.blendDisplayOnB = blendDisplayOnB;
            this.redMaskOnA = redMaskOnA;
            this.redMaskOnB = redMaskOnB;
        }
    }

    /**
     * 画像A・Bを読み込み、設定に応じてブレンドまたは元画像の PNG を outA / outB に保存する。
     *
     * @param pathA 画像Aのパス
     * @param pathB 画像Bのパス
     * @param outA  出力PNG（移行前側）
     * @param outB  出力PNG（移行後側）
     * @return リサイズ案内と差分ピクセル数など
     * @throws IOException 読み込み失敗など
     */
    public static MarkedPairResult writeMarkedPair(Path pathA, Path pathB, Path outA, Path outB) throws IOException {
        return writeMarkedPair(pathA, pathB, outA, outB,
                new DiffSettings(DEFAULT_RGB_SUM_THRESHOLD, DEFAULT_RED_ALPHA, MASK_EXPAND_RADIUS,
                        true, false, false, true));
    }

    public static MarkedPairResult writeMarkedPair(Path pathA, Path pathB, Path outA, Path outB, DiffSettings settings)
            throws IOException {
        BufferedImage rawA = ImageIO.read(pathA.toFile());
        BufferedImage rawB = ImageIO.read(pathB.toFile());
        if (rawA == null || rawB == null) {
            throw new IOException("画像を読み込めませんでした: "
                    + (rawA == null ? pathA : pathB));
        }
        BufferedImage a = toArgb(rawA);
        BufferedImage b = toArgb(rawB);
        String resizeNote = null;
        int aw = a.getWidth();
        int ah = a.getHeight();
        if (b.getWidth() != aw || b.getHeight() != ah) {
            b = scaleToSize(b, aw, ah);
            resizeNote = String.format(
                    "画像Bを画像Aの解像度 (%d×%d) に合わせてリサイズして比較しました（元: %d×%d）。",
                    aw, ah, rawB.getWidth(), rawB.getHeight());
        }
        int w = aw;
        int h = ah;
        int threshold = settings == null ? DEFAULT_RGB_SUM_THRESHOLD : settings.rgbSumThreshold;
        float redAlpha = settings == null ? DEFAULT_RED_ALPHA : settings.redAlpha;
        int expandRadius = settings == null ? MASK_EXPAND_RADIUS : settings.maskExpandRadius;
        boolean blendOnA = settings == null || settings.blendDisplayOnA;
        boolean blendOnB = settings == null || settings.blendDisplayOnB;
        boolean maskOnA = settings != null && settings.redMaskOnA;
        boolean maskOnB = settings != null && settings.redMaskOnB;
        threshold = Math.max(0, threshold);
        redAlpha = Math.max(0f, Math.min(1f, redAlpha));
        expandRadius = Math.max(0, expandRadius);

        boolean[] mask = buildDiffMask(a, b, w, h, threshold);
        boolean[] expandedMask = expandMask(mask, w, h, expandRadius);
        int rawDiffPixelCount = countTrue(mask);
        int expandedDiffPixelCount = countTrue(expandedMask);
        if (outA.getParent() != null) {
            Files.createDirectories(outA.getParent());
        }
        boolean needBlendA = blendOnA && !maskOnA;
        boolean needBlendB = blendOnB && !maskOnB;
        BufferedImage blended = (needBlendA || needBlendB)
                ? blendSemiTransparent(a, b, w, h, BLEND_WEIGHT_EACH, BLEND_WEIGHT_EACH)
                : null;
        BufferedImage outImgA = maskOnA
                ? applyRedOverlay(a, w, h, expandedMask, redAlpha)
                : (needBlendA ? blended : copyArgb(a, w, h));
        BufferedImage outImgB = maskOnB
                ? applyRedOverlay(b, w, h, expandedMask, redAlpha)
                : (needBlendB ? blended : copyArgb(b, w, h));
        ImageIO.write(outImgA, "png", outA.toFile());
        ImageIO.write(outImgB, "png", outB.toFile());
        return new MarkedPairResult(resizeNote, rawDiffPixelCount, expandedDiffPixelCount, w, h);
    }

    private static int countTrue(boolean[] mask) {
        int n = 0;
        for (boolean b : mask) {
            if (b) n++;
        }
        return n;
    }

    /**
     * 指定サイズへ描画スケールする（画像A基準の比較用）。
     */
    private static BufferedImage scaleToSize(BufferedImage src, int targetW, int targetH) {
        if (src.getWidth() == targetW && src.getHeight() == targetH) {
            return src;
        }
        BufferedImage dst = new BufferedImage(targetW, targetH, BufferedImage.TYPE_INT_ARGB);
        Graphics2D g = dst.createGraphics();
        try {
            g.setRenderingHint(RenderingHints.KEY_INTERPOLATION,
                    RenderingHints.VALUE_INTERPOLATION_BILINEAR);
            g.setRenderingHint(RenderingHints.KEY_RENDERING, RenderingHints.VALUE_RENDER_QUALITY);
            g.setRenderingHint(RenderingHints.KEY_ANTIALIASING, RenderingHints.VALUE_ANTIALIAS_ON);
            g.drawImage(src, 0, 0, targetW, targetH, null);
        } finally {
            g.dispose();
        }
        return dst;
    }

    private static BufferedImage copyArgb(BufferedImage src, int w, int h) {
        BufferedImage out = new BufferedImage(w, h, BufferedImage.TYPE_INT_ARGB);
        Graphics2D g = out.createGraphics();
        try {
            g.drawImage(src, 0, 0, null);
        } finally {
            g.dispose();
        }
        return out;
    }

    private static BufferedImage toArgb(BufferedImage src) {
        if (src.getType() == BufferedImage.TYPE_INT_ARGB) {
            return src;
        }
        BufferedImage dst = new BufferedImage(src.getWidth(), src.getHeight(), BufferedImage.TYPE_INT_ARGB);
        Graphics2D g = dst.createGraphics();
        try {
            g.drawImage(src, 0, 0, null);
        } finally {
            g.dispose();
        }
        return dst;
    }

    private static boolean[] buildDiffMask(BufferedImage a, BufferedImage b, int w, int h, int threshold) {
        int len = w * h;
        int[] pa = new int[len];
        int[] pb = new int[len];
        a.getRGB(0, 0, w, h, pa, 0, w);
        b.getRGB(0, 0, w, h, pb, 0, w);
        boolean[] mask = new boolean[len];
        for (int i = 0; i < len; i++) {
            int ca = pa[i];
            int cb = pb[i];
            int ra = (ca >> 16) & 0xff;
            int ga = (ca >> 8) & 0xff;
            int ba = ca & 0xff;
            int rb = (cb >> 16) & 0xff;
            int gb = (cb >> 8) & 0xff;
            int bb = cb & 0xff;
            int sum = Math.abs(ra - rb) + Math.abs(ga - gb) + Math.abs(ba - bb);
            mask[i] = sum > threshold;
        }
        return mask;
    }

    /**
     * 2枚をアルファブレンド（各ピクセルで weightA・weightB の比率で合成）。weight の合計が 1 のとき不透明度は維持しやすい。
     */
    private static BufferedImage blendSemiTransparent(
            BufferedImage a, BufferedImage b, int w, int h, float weightA, float weightB) {
        BufferedImage out = new BufferedImage(w, h, BufferedImage.TYPE_INT_ARGB);
        int len = w * h;
        int[] pa = new int[len];
        int[] pb = new int[len];
        a.getRGB(0, 0, w, h, pa, 0, w);
        b.getRGB(0, 0, w, h, pb, 0, w);
        int[] dst = new int[len];
        float sumW = weightA + weightB;
        if (sumW <= 0f) {
            sumW = 1f;
        }
        for (int i = 0; i < len; i++) {
            int ca = pa[i];
            int cb = pb[i];
            int aa = (ca >>> 24) & 0xff;
            int ab = (cb >>> 24) & 0xff;
            float ra = (ca >> 16) & 0xff;
            float ga = (ca >> 8) & 0xff;
            float ba = ca & 0xff;
            float rb = (cb >> 16) & 0xff;
            float gb = (cb >> 8) & 0xff;
            float bb = cb & 0xff;
            float outA = Math.min(255f, Math.round((aa * weightA + ab * weightB) / sumW));
            float outR = (ra * weightA + rb * weightB) / sumW;
            float outG = (ga * weightA + gb * weightB) / sumW;
            float outB = (ba * weightA + bb * weightB) / sumW;
            int ir = Math.min(255, Math.round(outR));
            int ig = Math.min(255, Math.round(outG));
            int ib = Math.min(255, Math.round(outB));
            int ia = Math.min(255, Math.round(outA));
            dst[i] = (ia << 24) | (ir << 16) | (ig << 8) | ib;
        }
        out.setRGB(0, 0, w, h, dst, 0, w);
        return out;
    }

    private static BufferedImage applyRedOverlay(
            BufferedImage baseArgb, int w, int h, boolean[] diffMask, float redAlpha) {
        BufferedImage out = new BufferedImage(w, h, BufferedImage.TYPE_INT_ARGB);
        int len = w * h;
        int[] src = new int[len];
        baseArgb.getRGB(0, 0, w, h, src, 0, w);
        float a = redAlpha;
        float inv = 1f - a;
        int[] dst = new int[len];
        for (int i = 0; i < len; i++) {
            if (!diffMask[i]) {
                dst[i] = src[i];
                continue;
            }
            int c = src[i];
            int r = (c >> 16) & 0xff;
            int g = (c >> 8) & 0xff;
            int b = c & 0xff;
            int alpha = (c >>> 24) & 0xff;
            int nr = Math.min(255, Math.round(r * inv + 255f * a));
            int ng = Math.min(255, Math.round(g * inv));
            int nb = Math.min(255, Math.round(b * inv));
            dst[i] = (alpha << 24) | (nr << 16) | (ng << 8) | nb;
        }
        out.setRGB(0, 0, w, h, dst, 0, w);
        return out;
    }

    private static boolean[] expandMask(boolean[] mask, int w, int h, int radius) {
        if (radius <= 0) {
            return mask;
        }
        boolean[] expanded = new boolean[w * h];
        for (int y = 0; y < h; y++) {
            for (int x = 0; x < w; x++) {
                int idx = y * w + x;
                if (!mask[idx]) {
                    continue;
                }
                int y0 = Math.max(0, y - radius);
                int y1 = Math.min(h - 1, y + radius);
                int x0 = Math.max(0, x - radius);
                int x1 = Math.min(w - 1, x + radius);
                for (int yy = y0; yy <= y1; yy++) {
                    for (int xx = x0; xx <= x1; xx++) {
                        expanded[yy * w + xx] = true;
                    }
                }
            }
        }
        return expanded;
    }
}

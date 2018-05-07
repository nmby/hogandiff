package xyz.hotchpotch.hogandiff.poi;

import java.io.File;
import java.io.InputStream;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.EnumSet;
import java.util.Objects;
import java.util.Set;
import java.util.stream.Collectors;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellAddress;

import xyz.hotchpotch.hogandiff.ApplicationException;
import xyz.hotchpotch.hogandiff.Context;
import xyz.hotchpotch.hogandiff.Context.Props;
import xyz.hotchpotch.hogandiff.common.BookType;
import xyz.hotchpotch.hogandiff.common.Pair;
import xyz.hotchpotch.hogandiff.common.Pair.Side;
import xyz.hotchpotch.hogandiff.excel.BResult;
import xyz.hotchpotch.hogandiff.excel.SResult;

/**
 * HSSF（.xls）形式およびXSSF（.xlsx, .xlsm）形式のExcelブックにPOIのユーザーモデルAPIを使用して色を付けるための
 * {@link BookPainter} の実装です。<br>
 * 
 * @author nmby
 * @since 0.3.1
 */
/*package*/ class BookPainterWithUserApi implements BookPainter {
    
    // [static members] ********************************************************
    
    /** このクラスがサポートするブック形式 */
    private static final Set<BookType> supported = EnumSet.of(
            BookType.XLS, BookType.XLSX, BookType.XLSM);
    
    /**
     * このクラスが指定されたファイルの形式をサポートするかを返します。<br>
     * 
     * @param file 検査対象のファイル
     * @return このクラスが指定されたファイルをサポートする場合は {@code true}
     * @throws NullPointerException {@code file} が {@code null} の場合
     */
    public static boolean isSupported(File file) {
        Objects.requireNonNull(file);
        return supported.stream()
                .map(BookType::extension)
                .anyMatch(file.getName()::endsWith);
    }
    
    /**
     * ペインターオブジェクトを生成して返します。<br>
     * 
     * @param context コンテキスト
     * @return 新しいペインター
     * @throws NullPointerException {@code context} が {@code null} の場合
     */
    public static BookPainterWithUserApi of(Context context) {
        Objects.requireNonNull(context, "context");
        return new BookPainterWithUserApi(context);
    }
    
    // [instance members] ******************************************************
    
    private final short redundantColor;
    private final short diffColor;
    
    private BookPainterWithUserApi(Context context) {
        assert context != null;
        
        redundantColor = context.get(Props.APP_REDUNDANT_COLOR);
        diffColor = context.get(Props.APP_DIFF_COLOR);
    }
    
    /**
     * {@inheritDoc}
     * 
     * @throws NullPointerException {@code src}, {@code dst}, {@code bResult}, {@code sides}
     *                              のいずれかが {@code null} の場合
     * @throws IllegalArgumentException {@code src} のファイル形式がこのクラスのサポート対象外の場合
     */
    @Override
    public void paintAndSave(
            File book,
            Path copy,
            BResult bResult,
            Side... sides)
            throws ApplicationException {
        
        Objects.requireNonNull(book, "book");
        Objects.requireNonNull(copy, "copy");
        Objects.requireNonNull(bResult, "bResult");
        Objects.requireNonNull(sides, "sides");
        if (!isSupported(book)) {
            throw new IllegalArgumentException(book.getPath());
        }
        
        // 1. 目的のファイルをコピーする。
        try {
            Files.copy(book.toPath(), copy);
            File newFile = copy.toFile();
            newFile.setReadable(true, false);
            newFile.setWritable(true, false);
        } catch (Exception e) {
            throw new ApplicationException("Excelファイルのコピーに失敗しました。\n" + copy.toString(), e);
        }
        
        // 2. コピーしたファイルをExcelブックとしてロードする。
        try (InputStream is = Files.newInputStream(copy);
                Workbook wb = WorkbookFactory.create(is)) {
            
            // 3. 色を付ける。
            POIUtils.clearColors(wb);
            bResult.sheetNamePairs.stream()
                    .filter(Pair::isPaired)
                    .forEach(p -> {
                        SResult sResult = bResult.sResults.get(p);
                        for (Pair.Side side : sides) {
                            Sheet sheet = wb.getSheet(p.get(side));
                            paintSheet(sheet, sResult.pieces.get(side));
                        }
                    });
            
            // 4. 着色したExcelブックをコピーしたファイルに上書き保存する。
            try (OutputStream os = Files.newOutputStream(copy)) {
                wb.write(os);
            }
            
        } catch (Exception e) {
            throw new ApplicationException("比較結果Excelブックの作成と保存に失敗しました。", e);
        } finally {
            System.gc();
        }
    }
    
    /**
     * 指定されたシートに比較結果の色を付けます。<br>
     * 
     * @param sheet 対象のシート
     * @param piece 対象シート上の差分箇所
     */
    private void paintSheet(Sheet sheet, SResult.Piece piece) {
        assert sheet != null;
        assert piece != null;
        
        POIUtils.paintRows(sheet, piece.redundantRows, redundantColor);
        POIUtils.paintColumns(sheet, piece.redundantColumns, redundantColor);
        Set<CellAddress> addresses = piece.diffCells.stream()
                .map(cell -> new CellAddress(cell.row(), cell.column()))
                .collect(Collectors.toSet());
        POIUtils.paintCells(sheet, addresses, diffColor);
    }
}

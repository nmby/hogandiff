package xyz.hotchpotch.hogandiff.excel;

import java.io.File;
import java.util.Objects;
import java.util.Set;

import xyz.hotchpotch.hogandiff.ApplicationException;
import xyz.hotchpotch.hogandiff.excel.hssf.HSSFSheetLoaderWithEventApi;
import xyz.hotchpotch.hogandiff.excel.xssf.XSSFSheetLoaderWithEventApi;

/**
 * Excelブックからシートデータを読み込みローダーを表します。<br>
 * これは、{@link #loadSheet(File, String)} を関数メソッドに持つ関数型インタフェースです。<br>
 * 
 * @author nmby
 * @since 0.1.0
 */
@FunctionalInterface
public interface SheetLoader {
    
    // [static members] ********************************************************
    
    /**
     * 指定されたExcelブックに適したローダーを返します。<br>
     * 
     * @param book 対象のExcelブック
     * @param extractCachedValue 数式セルから数式文字列ではなくキャッシュされた計算値を取得する場合は {@code true}
     * @return 指定されたExcelブックに適したローダー
     * @throws UnsupportedOperationException 指定されたExcelブックに適合するローダーが無い場合
     * @throws NullPointerException {@code book} が {@code null} の場合
     */
    public static SheetLoader of(File book, boolean extractCachedValue) {
        Objects.requireNonNull(book, "book");
        
        if (XSSFSheetLoaderWithEventApi.isSupported(book)) {
            return XSSFSheetLoaderWithEventApi.of(extractCachedValue);
        }
        if (HSSFSheetLoaderWithEventApi.isSupported(book) && extractCachedValue) {
            return HSSFSheetLoaderWithEventApi.of(extractCachedValue);
        }
        if (SheetLoaderWithUserApi.isSupported(book)) {
            return SheetLoaderWithUserApi.of(extractCachedValue);
        }
        throw new UnsupportedOperationException(book.getPath());
    }
    
    // [instance members] ******************************************************
    
    /**
     * シートデータを読み込んでセルデータのセットとして返します。<br>
     * 
     * @param book 対象のExcelブック
     * @param sheetName 対象のシート名
     * @return シートに含まれるセルデータのセット
     * @throws ApplicationException 処理に失敗した場合
     */
    Set<CellReplica> loadSheet(File book, String sheetName) throws ApplicationException;
}

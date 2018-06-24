package xyz.hotchpotch.hogandiff.excel;

import java.io.File;
import java.nio.file.Path;
import java.util.List;
import java.util.Map;
import java.util.Objects;

import xyz.hotchpotch.hogandiff.ApplicationException;
import xyz.hotchpotch.hogandiff.Context;
import xyz.hotchpotch.hogandiff.diff.excel.SResult;
import xyz.hotchpotch.hogandiff.excel.xssf.XSSFBookPainterWithStaxApi;

/**
 * Excelブックの差分箇所に色を付けるペインターを表します。<br>
 * これは、{@link #paintAndSave(File, Path, List)} を関数メソッドに持つ関数型インタフェースです。<br>
 * 
 * @author nmby
 * @since 0.3.2
 */
@FunctionalInterface
public interface BookPainter {
    
    // [static members] ********************************************************
    
    /**
     * 指定されたExcelブックに適したペインターを返します。<br>
     * 
     * @param book 対象のExcelブック
     * @param context コンテキスト
     * @return 指定されたExcelブックに適したペインター
     * @throws UnsupportedOperationException 指定されたExcelブックに適合するペインターが無い場合
     * @throws NullPointerException {@code book}, {@code context} のいずれかが {@code null} の場合
     */
    public static BookPainter of(File book, Context context) {
        Objects.requireNonNull(book, "book");
        Objects.requireNonNull(context, "context");
        
        if (XSSFBookPainterWithStaxApi.isSupported(book)) {
            return XSSFBookPainterWithStaxApi.of(context);
        }
        if (BookPainterWithUserApi.isSupported(book)) {
            return BookPainterWithUserApi.of(context);
        }
        throw new UnsupportedOperationException(book.getPath());
    }
    
    // [instance members] ******************************************************
    
    /**
     * 指定されたExcelに色を付け、指定されたパスに保存します。<br>
     * 
     * @param book 処理対象Excelブックのファイルパス
     * @param copy 保存先ファイルパス
     * @param results シート名とその差分箇所のリスト
     * @throws ApplicationException 処理に失敗した場合
     */
    void paintAndSave(
            File book,
            Path copy,
            List<Map.Entry<String, SResult.Piece>> results)
            throws ApplicationException;
}

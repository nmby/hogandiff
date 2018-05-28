package xyz.hotchpotch.hogandiff.excel;

import java.io.File;
import java.util.List;

import xyz.hotchpotch.hogandiff.ApplicationException;
import xyz.hotchpotch.hogandiff.excel.hssf.HSSFSheetListerWithEventApi;
import xyz.hotchpotch.hogandiff.excel.xssf.XSSFSheetListerWithEventApi;

/**
 * Excelブックに含まれるシート名の一覧を返すリスターを表します。<br>
 * これは、{@link #getSheetNames(File)} を関数メソッドに持つ関数型インタフェースです。<br>
 * 
 * @author nmby
 * @since 0.3.2
 */
@FunctionalInterface
public interface SheetLister {
    
    // [static members] ********************************************************
    
    /**
     * 指定されたExcelブックに適したリスターを返します。<br>
     * 
     * @param book 対象のExcelブック
     * @return 指定されたExcelブックに適したリスター
     * @throws UnsupportedOperationException 指定されたExcelブックに適合するリスターが無い場合
     */
    public static SheetLister of(File book) {
        if (XSSFSheetListerWithEventApi.isSupported(book)) {
            return XSSFSheetListerWithEventApi.of();
        }
        if (HSSFSheetListerWithEventApi.isSupported(book)) {
            return HSSFSheetListerWithEventApi.of();
        }
        throw new UnsupportedOperationException(book.getPath());
    }
    
    // [instance members] ******************************************************
    
    /**
     * 指定されたExcelブックに含まれるシート名の一覧を返します。
     * 但し、グラフシートは対象外とし、戻り値には含まれません。<br>
     * 
     * @param book 対象のExcelブック
     * @return シート名の一覧
     * @throws ApplicationException 処理に失敗した場合
     */
    List<String> getSheetNames(File book) throws ApplicationException;
}

package xyz.hotchpotch.hogandiff.excel.xssf;

import java.io.File;
import java.util.EnumSet;
import java.util.List;
import java.util.Objects;
import java.util.Set;

import xyz.hotchpotch.hogandiff.ApplicationException;
import xyz.hotchpotch.hogandiff.excel.BookType;
import xyz.hotchpotch.hogandiff.excel.SheetLister;

/**
 * XSSF（.xlsx, .xlsm）形式のExcelブックからPOIのイベントモデルAPIを使用してシート名の一覧を取得するための
 * {@link SheetLister} の実装です。<br>
 * 
 * @author nmby
 * @since 0.3.2
 */
public class XSSFSheetListerWithEventApi implements SheetLister {
    
    // [static members] ********************************************************
    
    /** このクラスがサポートするブック形式 */
    private static final Set<BookType> supported = EnumSet.of(BookType.XLSX, BookType.XLSM);
    
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
     * リスターオブジェクトを生成して返します。<br>
     * 
     * @return 新しいリスター
     */
    public static XSSFSheetListerWithEventApi of() {
        return new XSSFSheetListerWithEventApi();
    }
    
    // [instance members] ******************************************************
    
    private XSSFSheetListerWithEventApi() {
    }
    
    /**
     * {@inheritDoc}
     * 
     * @throws NullPointerException {@code book} が {@code null} の場合
     * @throws IllegalArgumentException {@code book} のブック形式をこのクラスがサポートしない場合
     */
    @Override
    public List<String> getSheetNames(File book) throws ApplicationException {
        Objects.requireNonNull(book, "book");
        if (!isSupported(book)) {
            throw new IllegalArgumentException(book.getName());
        }
        
        XSSFSheetEntryManager manager = XSSFSheetEntryManager.generate(book.toPath());
        return manager.getWorksheetNames();
    }
}

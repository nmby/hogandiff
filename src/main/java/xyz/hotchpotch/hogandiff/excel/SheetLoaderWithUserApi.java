package xyz.hotchpotch.hogandiff.excel;

import java.io.File;
import java.util.EnumSet;
import java.util.NoSuchElementException;
import java.util.Objects;
import java.util.Set;
import java.util.stream.Collectors;
import java.util.stream.StreamSupport;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import xyz.hotchpotch.hogandiff.ApplicationException;

/**
 * HSSF（.xls）形式およびXSSF（.xlsx, .xlsm）形式のExcelブックからPOIのユーザーモデルAPIを使用してシートデータを読み込むための
 * {@link SheetLoader} の実装です。<br>
 * 
 * @author nmby
 * @since 0.2.0
 */
/*package*/class SheetLoaderWithUserApi implements SheetLoader {
    
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
     * ローダーオブジェクトを生成して返します。<br>
     * 
     * @param extractCachedValue 数式セルから数式文字列ではなくキャッシュされた計算値を取得する場合は {@code true}
     * @return 新しいローダー
     */
    public static SheetLoaderWithUserApi of(boolean extractCachedValue) {
        return new SheetLoaderWithUserApi(extractCachedValue);
    }
    
    // [instance members] ******************************************************
    
    private final boolean extractCachedValue;
    
    private SheetLoaderWithUserApi(boolean extractCachedValue) {
        this.extractCachedValue = extractCachedValue;
    }
    
    /**
     * {@inheritDoc}
     * 
     * @throws NullPointerException {@code book}, {code sheetName} のいずれかが {@code null} の場合
     * @throws IllegalArgumentException {@code book} のブック形式をこのクラスがサポートしない場合
     * @throws NoSuchElementException {@code sheetName} に該当するシートが存在しない場合
     */
    @Override
    public Set<CellReplica> loadSheet(File book, String sheetName) throws ApplicationException {
        Objects.requireNonNull(book, "book");
        Objects.requireNonNull(sheetName, "sheetName");
        if (!isSupported(book)) {
            throw new IllegalArgumentException(book.getName());
        }
        
        try (Workbook wb = WorkbookFactory.create(book)) {
            Sheet sheet = wb.getSheet(sheetName);
            if (sheet == null) {
                throw new NoSuchElementException(sheetName);
            }
            
            return StreamSupport.stream(sheet.spliterator(), true)
                    .flatMap(row -> StreamSupport.stream(row.spliterator(), false))
                    .filter(cell -> cell.getCellTypeEnum() != CellType.BLANK)
                    .map(cell -> CellReplica.of(
                            cell.getRowIndex(),
                            cell.getColumnIndex(),
                            ExcelUtils.getValue(cell, extractCachedValue)))
                    .filter(cell -> !"".equals(cell.value()))
                    .collect(Collectors.toSet());
            
        } catch (NoSuchElementException e) {
            throw e;
        } catch (Exception e) {
            String msg = String.format("シートの読み込みに失敗しました。book:%s, sheetName:%s",
                    book.getPath(), sheetName);
            throw new ApplicationException(msg, e);
        }
    }
}

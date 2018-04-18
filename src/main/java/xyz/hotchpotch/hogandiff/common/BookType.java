package xyz.hotchpotch.hogandiff.common;

import java.io.File;
import java.util.NoSuchElementException;
import java.util.Objects;
import java.util.stream.Stream;

/**
 * Excelブックの形式を表す列挙型です。<br>
 * 
 * @author nmby
 * @since 0.1.0
 */
public enum BookType {
    
    // [static members] ********************************************************
    
    /** .xlsx 形式 */
    XLSX(".xlsx"),
    
    /** .xlsm 形式 */
    XLSM(".xlsm"),
    
    /** .xlsb 形式 */
    XLSB(".xlsb"),
    
    /** .xls 形式 */
    XLS(".xls");
    
    /**
     * 指定されたファイルに対応する {@link BookType} インスタンスを返します。<br>
     * 
     * @param file 対象ファイル
     * @return 指定されたファイルに対応する {@link BookType} インスタンス
     * @throws NullPointerException {@code file} が {@code null} の場合
     * @throws NoSuchElementException 指定されたファイルに対応する {@link BookType} オブジェクトが存在しない場合
     */
    public static BookType of(File file) {
        Objects.requireNonNull(file, "file");
        
        String fileName = file.getName();
        return Stream.of(values())
                .filter(type -> fileName.endsWith(type.extension))
                .findFirst()
                .orElseThrow(() -> new NoSuchElementException(file.getName()));
    }
    
    // [instance members] ******************************************************
    
    private final String extension;
    
    private BookType(String extension) {
        this.extension = extension;
    }
    
    /**
     * このブック形式の拡張子（{@code ".xlsx"} など）を返します。<br>
     * 
     * @return このブック形式の拡張子（{@code ".xlsx"} など）
     */
    public String extension() {
        return extension;
    }
}

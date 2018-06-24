package xyz.hotchpotch.hogandiff.excel;

import java.util.Objects;

import xyz.hotchpotch.hogandiff.common.Pair;

/**
 * セルを表す簡易な不変クラスです。<br>
 * 
 * @author nmby
 * @since 0.3.2
 */
public class CellReplica {
    
    // [static members] ********************************************************
    
    private static final int NUM = 'Z' - 'A' + 1;
    
    /**
     * {@link CellReplica} オブジェクトを生成します。<br>
     * 
     * @param row 行インデックス（0開始）
     * @param column 列インデックス（0開始）
     * @param value セルの値（空セルの場合は {@code null} ではなく {@code ""}）
     * @return 新しい {@link CellReplica} オブジェクト
     * @throws NullPointerException {@code value} が {@code null} の場合
     * @throws IllegalArgumentException {@code row}, {@code column} のいずれかが {@code 0} 未満の場合
     */
    public static CellReplica of(int row, int column, String value) {
        Objects.requireNonNull(value, "value");
        if (row < 0 || column < 0) {
            throw new IllegalArgumentException(String.format("(%d, %d)", row, column));
        }
        return new CellReplica(row, column, value);
    }
    
    /**
     * {@link CellReplica} オブジェクトを生成します。<br>
     * 
     * @param address セルのアドレス（"A1"..）
     * @param value セルの値（空セルの場合は {@code null} ではなく {@code ""}）
     * @return 新しい {@link CellReplica} オブジェクト
     * @throws NullPointerException {@code address}, {@code value} のいずれかが {@code null} の場合
     * @throws IllegalArgumentException {@code address} の値が不正な場合
     */
    public static CellReplica of(String address, String value) {
        Objects.requireNonNull(address, "address");
        Objects.requireNonNull(value, "value");
        
        int i = 0;
        while (i < address.length()) {
            char c = address.charAt(i);
            if (c < 'A' || 'Z' < c) {
                break;
            }
            i++;
        }
        
        try {
            String columnName = address.substring(0, i);
            String rowName = address.substring(i);
            int row = Integer.parseInt(rowName) - 1;
            int column = getColumnIdx(columnName);
            return of(row, column, value);
            
        } catch (RuntimeException e) {
            throw new IllegalArgumentException("address: " + address);
        }
    }
    
    /**
     * 列名（"A"..）を列インデックス（0..）に変換して返します。<br>
     * 
     * @param columnName 列名（"A"..）
     * @return 列インデックス（0..）
     * @throws NullPointerException {@code columnName} が {@code null} の場合
     * @throws IllegalArgumentException {@code columnName} の値が不正な場合
     */
    private static int getColumnIdx(String columnName) {
        Objects.requireNonNull(columnName, "columnName");
        if (columnName.isEmpty()) {
            throw new IllegalArgumentException("columnName: " + columnName);
        }
        
        try {
            int idx = 0;
            for (int i = 0; i < columnName.length(); i++) {
                idx *= NUM;
                idx += (columnName.charAt(i) - 'A' + 1);
            }
            return idx - 1;
            
        } catch (RuntimeException e) {
            throw new IllegalArgumentException("columnName: " + columnName);
        }
    }
    
    /**
     * 指定されたセルアドレス（"A1"..）に対応する行インデックス（0開始）と列インデックス（0開始）のペアを返します。<br>
     * 
     * @param address セルのアドレス（"A1"..）
     * @return 行インデックス（0開始）と列インデックス（0開始）のペア
     * @throws NullPointerException {@code address} が {@code null} の場合
     * @throws IllegalArgumentException {@code address} の値が不正な場合
     * 
     * @since 0.4.0
     */
    public static Pair<Integer> getIndex(String address) {
        Objects.requireNonNull(address, "address");
        
        int i = 0;
        while (i < address.length()) {
            char c = address.charAt(i);
            if (c < 'A' || 'Z' < c) {
                break;
            }
            i++;
        }
        
        try {
            String columnName = address.substring(0, i);
            String rowName = address.substring(i);
            int row = Integer.parseInt(rowName) - 1;
            int column = getColumnIdx(columnName);
            return Pair.of(row, column);
            
        } catch (RuntimeException e) {
            throw new IllegalArgumentException("address: " + address);
        }
    }
    
    /**
     * 列インデックス（0..）を列名（"A"..）に変換して返します。<br>
     * 
     * @param columnIdx 列インデックス（0..）
     * @return 列名（"A"..）
     * @throws IllegalArgumentException {@code columnIdx} が {@code 0} 未満の場合
     */
    public static String getColumnName(int columnIdx) {
        if (columnIdx < 0) {
            throw new IllegalArgumentException("columnIdx: " + columnIdx);
        }
        
        int div = columnIdx + 1;
        StringBuilder str = new StringBuilder();
        
        do {
            int mod = div % NUM;
            div /= NUM;
            
            if (mod == 0) {
                mod = NUM;
                div--;
            }
            str.append((char) ('A' + mod - 1));
        } while (0 < div);
        
        return str.reverse().toString();
    }
    
    /**
     * 行インデックス（0..）と列インデックス（0..）をセルアドレス（"A1"..）に変換して返します。<br>
     * 
     * @param rowIdx 行インデックス（0..）
     * @param columnIdx 列インデックス（0..）
     * @return セルアドレス（"A1"..）
     * @throws IllegalArgumentException {@code rowIdx}, {@code columnIdx} のいずれかが {@code 0} 未満の場合
     */
    public static String getAddress(int rowIdx, int columnIdx) {
        if (rowIdx < 0 || columnIdx < 0) {
            throw new IllegalArgumentException(
                    String.format("rowIdx: %d, columnIdx: %d", rowIdx, columnIdx));
        }
        return String.format("%s%d", getColumnName(columnIdx), rowIdx + 1);
    }
    
    // [instance members] ******************************************************
    
    private final int row;
    private final int column;
    private final String value;
    
    private CellReplica(int row, int column, String value) {
        assert 0 <= row;
        assert 0 <= column;
        assert value != null;
        
        this.row = row;
        this.column = column;
        this.value = value;
    }
    
    // 不変なフィールドは getter を設けずに直接公開してしまってもよいのだが、
    // https://www.ibm.com/developerworks/jp/java/library/j-ft4/index.html
    // ラムダ式と相性が悪いので getter を設けることにする。
    
    /**
     * このセルの行インデックス（0開始）を返します。<br>
     * 
     * @return 行インデックス（0開始）
     */
    public int row() {
        return row;
    }
    
    /**
     * このセルの列インデックス（0開始）を返します。<br>
     * 
     * @return 列インデックス（0開始）
     */
    public int column() {
        return column;
    }
    
    /**
     * このセルのアドレス（"A1"..）を返します。<br>
     * 
     * @return セルのアドレス（"A1"..）
     */
    public String address() {
        return getAddress(row, column);
    }
    
    /**
     * このセルの値（空セルの場合は {@code null} ではなく {@code ""}）を返します。<br>
     * 
     * @return セルの値（空セルの場合は {@code null} ではなく {@code ""}）
     */
    public String value() {
        return value;
    }
    
    @Override
    public boolean equals(Object o) {
        if (o instanceof CellReplica) {
            CellReplica c = (CellReplica) o;
            return row == c.row && column == c.column && value.equals(c.value);
        }
        return false;
    }
    
    @Override
    public int hashCode() {
        return Objects.hash(row, column, value);
    }
    
    @Override
    public String toString() {
        return String.format("%s [%s]", address(), value);
    }
}

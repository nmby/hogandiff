package xyz.hotchpotch.hogandiff.excel;

import java.util.ArrayList;
import java.util.Collections;
import java.util.List;
import java.util.Objects;
import java.util.Optional;

import xyz.hotchpotch.hogandiff.common.CellReplica;
import xyz.hotchpotch.hogandiff.common.Pair;

/**
 * Excelシート同士の比較結果を表す不変クラスです。<br>
 * 
 * @author nmby
 * @since 0.1.0
 */
public class SResult {
    
    // [static members] ********************************************************
    
    private static final String BR = System.lineSeparator();
    
    /**
     * {@link SResult} オブジェクトを生成して返します。<br>
     * 
     * @param redundantRowsA シートA上の余剰行インデックスのリスト（行の挿入／削除を考慮しなかった場合は {@code null}）
     * @param redundantRowsB シートB上の余剰行インデックスのリスト（行の挿入／削除を考慮しなかった場合は {@code null}）
     * @param redundantColumnsA シートA上の余剰列インデックスのリスト（列の挿入／削除を考慮しなかった場合は {@code null}）
     * @param redundantColumnsB シートB上の余剰列インデックスのリスト（列の挿入／削除を考慮しなかった場合は {@code null}）
     * @param diffCells 差分セルを表すペアのリスト
     * @return 新しい {@link SResult} オブジェクト
     * @throws NullPointerException {@code diffCells} が {@code null} の場合
     * @throws IllegalArgumentException
     *      {@code redundantRowsA}, {@code redundantRowsB} の一方のみが {@code null} の場合や、
     *      {@code redundantColumnsA}, {@code redundantColumnsB} の一方のみが {@code null} の場合
     */
    public static SResult of(
            List<Integer> redundantRowsA,
            List<Integer> redundantRowsB,
            List<Integer> redundantColumnsA,
            List<Integer> redundantColumnsB,
            List<Pair<CellReplica>> diffCells) {
        
        Objects.requireNonNull(diffCells, "diffCells");
        if ((redundantRowsA == null) != (redundantRowsB == null)
                || (redundantColumnsA == null) != (redundantColumnsB == null)) {
            throw new IllegalArgumentException();
        }
        
        return new SResult(
                redundantRowsA,
                redundantRowsB,
                redundantColumnsA,
                redundantColumnsB,
                diffCells);
    }
    
    // [instance members] ******************************************************
    
    // 不変なフィールドは getter を設けずに直接公開してしまう。
    // https://www.ibm.com/developerworks/jp/java/library/j-ft4/index.html
    
    /** 余剰行インデックス（0開始）のリスト。比較において行の余剰/欠損を考慮しなかった場合は {@link Optional#empty()} */
    public final Optional<Pair<List<Integer>>> redundantRows;
    
    /** 余剰列インデックス（0開始）のリスト。比較において列の余剰/欠損を考慮しなかった場合は {@link Optional#empty()} */
    public final Optional<Pair<List<Integer>>> redundantColumns;
    
    /** 差分セルのリスト。差分なしの場合は長さ0のリスト */
    public final List<Pair<CellReplica>> diffCells;
    
    private SResult(
            List<Integer> redundantRowsA,
            List<Integer> redundantRowsB,
            List<Integer> redundantColumnsA,
            List<Integer> redundantColumnsB,
            List<Pair<CellReplica>> diffCells) {
        
        assert (redundantRowsA == null) == (redundantRowsB == null);
        assert (redundantColumnsA == null) == (redundantColumnsB == null);
        assert diffCells != null;
        
        this.redundantRows = (redundantRowsA == null)
                ? Optional.empty()
                : Optional.of(Pair.of(
                        // このオブジェクトを不変にするために防御的コピーをしたうえで変更不可コレクションでマップする。
                        Collections.unmodifiableList(new ArrayList<>(redundantRowsA)),
                        Collections.unmodifiableList(new ArrayList<>(redundantRowsB))));
        this.redundantColumns = (redundantColumnsA == null)
                ? Optional.empty()
                : Optional.of(Pair.of(
                        Collections.unmodifiableList(new ArrayList<>(redundantColumnsA)),
                        Collections.unmodifiableList(new ArrayList<>(redundantColumnsB))));
        this.diffCells = Collections.unmodifiableList(new ArrayList<>(diffCells));
    }
    
    /**
     * 比較結果のサマリを返します。<br>
     * 
     * @return 比較結果のサマリ
     */
    public String getSummary() {
        StringBuilder str = new StringBuilder();
        
        redundantRows.ifPresent(p -> str.append(String.format(
                "\t余剰行 : シートA - %s, シートB - %s",
                p.a().get().isEmpty() ? "（なし）" : p.a().get().size() + "行",
                p.b().get().isEmpty() ? "（なし）" : p.b().get().size() + "行"))
                .append(BR));
        
        redundantColumns.ifPresent(p -> str.append(String.format(
                "\t余剰列 : シートA - %s, シートB - %s",
                p.a().get().isEmpty() ? "（なし）" : p.a().get().size() + "列",
                p.b().get().isEmpty() ? "（なし）" : p.b().get().size() + "列"))
                .append(BR));
        
        str.append(String.format("\t差分セル : %s",
                diffCells.isEmpty() ? "（なし）" : "各シート" + diffCells.size() + "セル"))
                .append(BR);
        
        return str.toString();
    }
    
    /**
     * 比較結果の詳細を返します。<br>
     * 
     * @return 比較結果の詳細
     */
    public String getDetail() {
        StringBuilder str = new StringBuilder();
        
        redundantRows.ifPresent(p -> {
            str.append(BR).append("\tシートA上の余剰行 :").append(BR);
            if (p.a().get().isEmpty()) {
                str.append("\t\t（なし）").append(BR);
            } else {
                p.a().get().forEach(i -> str.append("\t\t行").append(i + 1).append(BR));
            }
            str.append(BR).append("\tシートB上の余剰行 :").append(BR);
            if (p.b().get().isEmpty()) {
                str.append("\t\t（なし）").append(BR);
            } else {
                p.b().get().forEach(i -> str.append("\t\t行").append(i + 1).append(BR));
            }
            str.append(BR);
        });
        redundantColumns.ifPresent(p -> {
            str.append("\tシートA上の余剰列 :").append(BR);
            if (p.a().get().isEmpty()) {
                str.append("\t\t（なし）").append(BR);
            } else {
                p.a().get().forEach(j -> str.append(
                        String.format("\t\t%s列", CellReplica.getColumnName(j))).append(BR));
            }
            str.append(BR).append("\tシートB上の余剰列 :").append(BR);
            if (p.b().get().isEmpty()) {
                str.append("\t\t（なし）").append(BR);
            } else {
                p.b().get().forEach(j -> str.append(
                        String.format("\t\t%s列", CellReplica.getColumnName(j))).append(BR));
            }
            str.append(BR);
        });
        str.append("\t差分セル :");
        if (diffCells.isEmpty()) {
            str.append(BR).append("\t\t（なし）").append(BR);
        } else {
            diffCells.forEach(p -> {
                str.append(BR);
                str.append("\t\tセルA : ").append(p.a().get()).append(BR);
                str.append("\t\tセルB : ").append(p.b().get()).append(BR);
            });
        }
        
        return str.toString();
    }
}

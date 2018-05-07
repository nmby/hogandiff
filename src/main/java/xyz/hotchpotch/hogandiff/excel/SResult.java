package xyz.hotchpotch.hogandiff.excel;

import java.util.ArrayList;
import java.util.Collections;
import java.util.Iterator;
import java.util.List;
import java.util.Objects;
import java.util.stream.Collectors;

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
    
    /**
     * 2つのExcelシートの比較結果のうち、一方のシートの差分箇所を表す不変クラスです。<br>
     * 
     * @author nmby
     * @since 0.3.1
     */
    public static class Piece {
        
        // [static members] ----------------------------------------------------
        
        // [instance members] --------------------------------------------------
        
        // 不変なフィールドは getter を設けずに直接公開してしまう。
        // https://www.ibm.com/developerworks/jp/java/library/j-ft4/index.html
        
        /** このオブジェクトの親 {@link SResult} オブジェクト */
        public final SResult parent;
        
        /** 余剰行インデックス（0開始）のリスト */
        public final List<Integer> redundantRows;
        
        /** 余剰列インデックス（0開始）のリスト */
        public final List<Integer> redundantColumns;
        
        /** 差分セルのリスト */
        public final List<CellReplica> diffCells;
        
        private Piece(
                SResult parent,
                List<Integer> redundantRows,
                List<Integer> redundantColumns,
                List<CellReplica> diffCells) {
            
            assert parent != null;
            assert redundantRows != null;
            assert redundantColumns != null;
            assert diffCells != null;
            
            this.parent = parent;
            // このオブジェクトを不変にするために防御的コピーをしたうえで変更不可コレクションでマップする。
            this.redundantRows = Collections.unmodifiableList(new ArrayList<>(redundantRows));
            this.redundantColumns = Collections.unmodifiableList(new ArrayList<>(redundantColumns));
            this.diffCells = Collections.unmodifiableList(new ArrayList<>(diffCells));
        }
    }
    
    private static final String BR = System.lineSeparator();
    
    /**
     * {@link SResult} オブジェクトを生成して返します。<br>
     * 
     * @param considerRowGaps 比較において行の挿入／削除を考慮した場合は {@code true}
     * @param considerColumnGaps 比較において列の挿入／削除を考慮した場合は {@code true}
     * @param redundantRowsA シートA上の余剰行インデックス（0開始）のリスト
     * @param redundantRowsB シートB上の余剰行インデックス（0開始）のリスト
     * @param redundantColumnsA シートA上の余剰列インデックス（0開始）のリスト
     * @param redundantColumnsB シートB上の余剰列インデックス（0開始）のリスト
     * @param diffCells 差分セルを表すペアのリスト
     * @return 新しい {@link SResult} オブジェクト
     * @throws NullPointerException
     *      {@code redundantRowsA}, {@code redundantRowsA}, 
     *      {@code redundantColumnsA}, {@code redundantColumnsA}, 
     *      {@code diffCells} のいずれかが {@code null} の場合
     * @throws IllegalArgumentException
     *      {@code considerRowGaps} が {@code false} であるにも関わらず
     *      {@code redundantRowsA}, {@code redundantRowsB} のいずれかが {@code empty()} ではない場合や、
     *      {@code considerColumnGaps} が {@code false} であるにも関わらず
     *      {@code redundantColumnsA}, {@code redundantColumnsB} のいずれかが {@code empty()} ではない場合
     */
    public static SResult of(
            boolean considerRowGaps,
            boolean considerColumnGaps,
            List<Integer> redundantRowsA,
            List<Integer> redundantRowsB,
            List<Integer> redundantColumnsA,
            List<Integer> redundantColumnsB,
            List<Pair<CellReplica>> diffCells) {
        
        Objects.requireNonNull(redundantRowsA, "redundantRowsA");
        Objects.requireNonNull(redundantRowsB, "redundantRowsB");
        Objects.requireNonNull(redundantColumnsA, "redundantColumnsA");
        Objects.requireNonNull(redundantColumnsB, "redundantColumnsB");
        Objects.requireNonNull(diffCells, "diffCells");
        if ((!considerRowGaps && (!redundantRowsA.isEmpty() || !redundantRowsB.isEmpty()))
                || (!considerColumnGaps && (!redundantColumnsA.isEmpty() || !redundantColumnsB.isEmpty()))) {
            throw new IllegalArgumentException();
        }
        
        return new SResult(
                considerRowGaps,
                considerColumnGaps,
                redundantRowsA,
                redundantRowsB,
                redundantColumnsA,
                redundantColumnsB,
                diffCells);
    }
    
    // [instance members] ******************************************************
    
    // 不変なフィールドは getter を設けずに直接公開してしまう。
    // https://www.ibm.com/developerworks/jp/java/library/j-ft4/index.html
    
    /** 各シートの差分箇所 */
    public final Pair<Piece> pieces;
    
    private final boolean considerRowGaps;
    private final boolean considerColumnGaps;
    
    private SResult(
            boolean considerRowGaps,
            boolean considerColumnGaps,
            List<Integer> redundantRowsA,
            List<Integer> redundantRowsB,
            List<Integer> redundantColumnsA,
            List<Integer> redundantColumnsB,
            List<Pair<CellReplica>> diffCells) {
        
        assert redundantRowsA != null;
        assert redundantRowsB != null;
        assert redundantColumnsA != null;
        assert redundantColumnsB != null;
        assert diffCells != null;
        
        this.considerRowGaps = considerRowGaps;
        this.considerColumnGaps = considerColumnGaps;
        
        this.pieces = Pair.of(
                new Piece(
                        this,
                        redundantRowsA,
                        redundantColumnsA,
                        diffCells.stream().map(Pair::a).collect(Collectors.toList())),
                new Piece(
                        this,
                        redundantRowsB,
                        redundantColumnsB,
                        diffCells.stream().map(Pair::b).collect(Collectors.toList())));
    }
    
    /**
     * 比較結果のサマリを返します。<br>
     * 
     * @return 比較結果のサマリ
     */
    public String getSummary() {
        StringBuilder str = new StringBuilder();
        
        if (considerRowGaps) {
            List<Integer> rowsA = pieces.a().redundantRows;
            List<Integer> rowsB = pieces.b().redundantRows;
            str.append(String.format(
                    "\t余剰行 : シートA - %s, シートB - %s",
                    rowsA.isEmpty() ? "（なし）" : rowsA.size() + "行",
                    rowsB.isEmpty() ? "（なし）" : rowsB.size() + "行"))
                    .append(BR);
        }
        
        if (considerColumnGaps) {
            List<Integer> columnsA = pieces.a().redundantColumns;
            List<Integer> columnsB = pieces.b().redundantColumns;
            str.append(String.format(
                    "\t余剰列 : シートA - %s, シートB - %s",
                    columnsA.isEmpty() ? "（なし）" : columnsA.size() + "列",
                    columnsB.isEmpty() ? "（なし）" : columnsB.size() + "列"))
                    .append(BR);
        }
        
        List<CellReplica> cells = pieces.a().diffCells;
        str.append(String.format("\t差分セル : %s",
                cells.isEmpty() ? "（なし）" : "各シート" + cells.size() + "セル"))
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
        
        if (considerRowGaps) {
            for (Pair.Side side : Pair.Side.values()) {
                List<Integer> rows = pieces.get(side).redundantRows;
                str.append(BR).append(String.format("\tシート%s上の余剰行 :", side)).append(BR);
                if (rows.isEmpty()) {
                    str.append("\t\t（なし）").append(BR);
                } else {
                    rows.forEach(i -> str.append("\t\t行").append(i + 1).append(BR));
                }
            }
        }
        
        if (considerColumnGaps) {
            for (Pair.Side side : Pair.Side.values()) {
                List<Integer> columns = pieces.get(side).redundantColumns;
                str.append(BR).append(String.format("\tシート%s上の余剰列 :", side)).append(BR);
                if (columns.isEmpty()) {
                    str.append("\t\t（なし）").append(BR);
                } else {
                    columns.forEach(j -> str.append(
                            String.format("\t\t%s列", CellReplica.getColumnName(j))).append(BR));
                }
            }
        }
        
        str.append(BR).append("\t差分セル :");
        if (pieces.a().diffCells.isEmpty()) {
            str.append(BR).append("\t\t（なし）").append(BR);
        } else {
            Iterator<CellReplica> itrA = pieces.a().diffCells.iterator();
            Iterator<CellReplica> itrB = pieces.b().diffCells.iterator();
            while (itrA.hasNext()) {
                CellReplica cellA = itrA.next();
                CellReplica cellB = itrB.next();
                str.append(BR);
                str.append("\t\tセルA : ").append(cellA).append(BR);
                str.append("\t\tセルB : ").append(cellB).append(BR);
            }
        }
        
        return str.toString();
    }
}

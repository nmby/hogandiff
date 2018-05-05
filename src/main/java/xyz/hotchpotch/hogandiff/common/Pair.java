package xyz.hotchpotch.hogandiff.common;

import java.util.EnumMap;
import java.util.Map;
import java.util.NoSuchElementException;
import java.util.Objects;
import java.util.Optional;
import java.util.function.BiConsumer;
import java.util.function.Consumer;
import java.util.function.Function;

/**
 * 同型の2つの要素のペアを表すクラスです。<br>
 * 要素オブジェクトが不変な場合、このクラスのオブジェクトも不変です。<br>
 * <br>
 * これは値ベースのクラスで、{@link Pair} のインスタンスに対してID依存操作（参照等価性 {@code ==}、IDハッシュ・コード、
 * 同期など）を使用すると、予期できない結果になる可能性があり、避けてください。<br>
 * 
 * @param <T> 要素の型
 * @since 0.1.0
 * @author nmby
 */
public class Pair<T> {
    
    // [static members] ********************************************************
    
    /**
     * ペアのどちら側かを表す列挙型です。<br>
     * 
     * @author nmby
     * @since 0.1.0
     */
    public static enum Side {
        
        // [static members] ----------------------------------------------------
        
        /** A-side */
        A,
        
        /** B-side */
        B;
        
        // [instance members] --------------------------------------------------
        
        /**
         * 自身と逆の側を返します。<br>
         * 
         * @return 自身とは逆の側
         */
        public Side opposite() {
            return this == A ? B : A;
        }
    }
    
    /**
     * 要素a, 要素bがともに {@code null} 以外の {@link Pair} オブジェクトを生成して返します。<br>
     * 
     * @param <T> 要素の型
     * @param a 要素a
     * @param b 要素b
     * @return 要素aと要素bのペアを表す {@code Pair} オブジェクト
     * @throws NullPointerException {@code a}, {@code b} のいずれかが {@code null} の場合
     */
    public static <T> Pair<T> of(T a, T b) {
        Objects.requireNonNull(a, "a");
        Objects.requireNonNull(b, "b");
        return new Pair<>(a, b);
    }
    
    /**
     * 要素a, 要素bの一方または双方が {@code null} の可能性がある {@link Pair} オブジェクトを生成して返します。<br>
     * 
     * @param <T> 要素の型
     * @param a 要素a
     * @param b 要素b
     * @return 要素aと要素bのペアを表す {@code Pair} オブジェクト
     */
    public static <T> Pair<T> ofNullable(T a, T b) {
        return new Pair<>(a, b);
    }
    
    /**
     * 要素aに対応する要素bが存在しないことを表す {@link Pair} オブジェクトを生成して返します。<br>
     * 
     * @param <T> 要素の型
     * @param a 要素a
     * @return 要素aに対応する要素bが存在しないことを表す {@code Pair} オブジェクト
     * @throws NullPointerException {@code a} が {@code null} の場合
     */
    public static <T> Pair<T> onlyA(T a) {
        Objects.requireNonNull(a, "a");
        return new Pair<>(a, null);
    }
    
    /**
     * 要素bに対応する要素aが存在しないことを表す {@link Pair} オブジェクトを生成して返します。<br>
     * 
     * @param <T> 要素の型
     * @param b 要素b
     * @return 要素bに対応する要素aが存在しないことを表す {@code Pair} オブジェクト
     * @throws NullPointerException {@code b} が {@code null} の場合
     */
    public static <T> Pair<T> onlyB(T b) {
        Objects.requireNonNull(b, "b");
        return new Pair<>(null, b);
    }
    
    /**
     * 要素a, 要素bがともに {@code null} である空の {@link Pair} オブジェクトを生成して返します。<br>
     * @param <T> 要素の型
     * @return 要素a, 要素bがともに {@code null} である空の {@link Pair} オブジェクト
     */
    public static <T> Pair<T> empty() {
        return new Pair<>(null, null);
    }
    
    /**
     * 要素a, 要素bのペアを表す {@link Pair} オブジェクトを生成して返します。<br>
     * 
     * @param <T> 要素の型
     * @param a 要素a
     * @param b 要素b
     * @return 要素aと要素bのペアを表す {@code Pair} オブジェクト
     * @throws NullPointerException {@code a}, {@code b} のいずれかが {@code null} の場合
     */
    public static <T> Pair<T> flatOf(Optional<? extends T> a, Optional<? extends T> b) {
        Objects.requireNonNull(a, "a");
        Objects.requireNonNull(b, "b");
        return new Pair<>(a.orElse(null), b.orElse(null));
    }
    
    // [instance members] ******************************************************
    
    private final Optional<T> a;
    private final Optional<T> b;
    
    private Pair(T a, T b) {
        this.a = Optional.ofNullable(a);
        this.b = Optional.ofNullable(b);
    }
    
    /**
     * 要素aに値が存在する場合は値を返し、それ以外の場合はNoSuchElementExceptionをスローします。<br>
     * 
     * @return 要素aの非null値
     * @throws NoSuchElementException 要素aが値を保持しない場合
     */
    public T a() {
        return a.get();
    }
    
    /**
     * 要素aに値が存在する場合は値を返し、それ以外の場合はotherを返します。<br>
     * 
     * @param other 要素aが値を保持しない場合に返す値（nullも可）
     * @return 要素aの値（存在する場合）、それ以外の場合はother
     */
    public T aOrElse(T other) {
        return a.orElse(other);
    }
    
    /**
     * 要素bに値が存在する場合は値を返し、それ以外の場合はNoSuchElementExceptionをスローします。<br>
     * 
     * @return 要素bの非null値
     * @throws NoSuchElementException 要素bが値を保持しない場合
     */
    public T b() {
        return b.get();
    }
    
    /**
     * 要素bに値が存在する場合は値を返し、それ以外の場合はotherを返します。<br>
     * 
     * @param other 要素bが値を保持しない場合に返す値（nullも可）
     * @return 要素bの値（存在する場合）、それ以外の場合はother
     */
    public T bOrElse(T other) {
        return b.orElse(other);
    }
    
    /**
     * 指定された側の要素が存在する場合は値を返し、それ以外の場合はNoSuchElementExceptionをスローします。<br>
     * 
     * @param side 要素の側
     * @return 指定された側の要素の非null値
     * @throws NoSuchElementException 指定された側の要素が値を保持しない場合
     * @throws NullPointerException {@code side} が {@code null} の場合
     */
    public T get(Side side) {
        Objects.requireNonNull(side, "side");
        return side == Side.A ? a.get() : b.get();
    }
    
    /**
     * 要素aを返します。<br>
     * 
     * @return 要素a
     */
    @Deprecated
    public Optional<T> a2() {
        return a;
    }
    
    /**
     * 要素bを返します。<br>
     * 
     * @return 要素b
     */
    @Deprecated
    public Optional<T> b2() {
        return b;
    }
    
    /**
     * 指定された側の要素を返します。<br>
     * 
     * @param side 要素の側
     * @return 指定された側の要素
     * @throws NullPointerException {@code side} が {@code null} の場合
     */
    @Deprecated
    public Optional<T> get2(Side side) {
        Objects.requireNonNull(side, "side");
        return side == Side.A ? a : b;
    }
    
    /**
     * 要素a, 要素bが格納されたマップを返します。<br>
     * 
     * @return 要素a, 要素bが格納されたマップ
     */
    @Deprecated
    public Map<Side, Optional<T>> items() {
        Map<Side, Optional<T>> items = new EnumMap<>(Side.class);
        items.put(Side.A, a);
        items.put(Side.B, b);
        return items;
    }
    
    /**
     * 要素a, 要素bがともに格納されているかを返します。<br>
     * 一方または双方が {@code null} の場合は {@code false} を返します。<br>
     * 
     * @return 要素a, 要素bがともに格納されている場合は {@code true}
     */
    public boolean isPaired() {
        return a.isPresent() && b.isPresent();
    }
    
    /**
     * 要素aのみが、{@code null} 以外の値であるかを返します。<br>
     * 
     * @return 要素aのみが、{@code null} 以外の値である場合は {@code true}
     */
    public boolean isOnlyA() {
        return a.isPresent() && !b.isPresent();
    }
    
    /**
     * 要素bのみが、{@code null} 以外の値であるかを返します。<br>
     * 
     * @return 要素bのみが、{@code null} 以外の値である場合は {@code true}
     */
    public boolean isOnlyB() {
        return !a.isPresent() && b.isPresent();
    }
    
    /**
     * 指定された側の要素のみが {@code null} 以外の値であるかを返します。<br>
     * 
     * @param side 要素の側
     * @return 指定された側の要素のみが {@code null} 以外の値である場合は {@code true}
     * @throws NullPointerException {@code side} が {@code null} の場合
     */
    public boolean isOnly(Side side) {
        Objects.requireNonNull(side, "side");
        return get2(side).isPresent() && !get2(side.opposite()).isPresent();
    }
    
    /**
     * 要素aと要素bがともに {@code null} であるかを返します。<br>
     * 
     * @return 要素aと要素bがともに {@code null} の場合は {@code true}
     */
    public boolean isEmpty() {
        return !a.isPresent() && !b.isPresent();
    }
    
    /**
     * 要素aが非null値を保持する場合は {@code true}, それ以外の場合は {@code false} を返します。<br>
     * 
     * @return 要素aが非null値を保持する場合は {@code true}
     */
    public boolean isPresentA() {
        return a.isPresent();
    }
    
    /**
     * 要素bが非null値を保持する場合は {@code true}, それ以外の場合は {@code false} を返します。<br>
     * 
     * @return 要素bが非null値を保持する場合は {@code true}
     */
    public boolean isPresentB() {
        return b.isPresent();
    }
    
    /**
     * {@code null} でない要素a, 要素bそれぞれに指定された関数を適用して得られた要素を保持する
     * {@link Pair} オブジェクトを返します。<br>
     * 
     * @param <U> 変換後の要素の型
     * @param mapper 変換関数
     * @return 変換後の要素を保持する {@link Pair} オブジェクト
     * @throws NullPointerException {@code mapper} が {@code null} の場合
     */
    public <U> Pair<U> map(Function<? super T, ? extends U> mapper) {
        Objects.requireNonNull(mapper, "mapper");
        return Pair.flatOf(a.map(mapper), b.map(mapper));
    }
    
    /**
     * {@code null} でない要素a, 要素bそれぞれに指定された関数を適用して得られた要素を保持する
     * {@link Pair} オブジェクトを返します。<br>
     * 
     * @param <U> 変換後の要素の型
     * @param mapper 変換関数
     * @return 変換後の要素を保持する {@link Pair} オブジェクト
     * @throws NullPointerException {@code mapper} が {@code null} の場合
     */
    public <U> Pair<U> flatMap(Function<? super T, Optional<U>> mapper) {
        Objects.requireNonNull(mapper, "mapper");
        return Pair.flatOf(a.flatMap(mapper), b.flatMap(mapper));
    }
    
    /**
     * {@code null} でない要素a, 要素bそれぞれに対して指定されたアクションを行います。<br>
     * 
     * @param action アクション
     * @throws NullPointerException {@code action} が {@code null} の場合
     */
    public void forEach(Consumer<? super T> action) {
        Objects.requireNonNull(action, "action");
        a.ifPresent(action);
        b.ifPresent(action);
    }
    
    /**
     * {@code null} でない要素a, 要素bそれぞれに対して指定されたアクションを行います。<br>
     * 
     * @param action アクション
     * @throws NullPointerException {@code action} が {@code null} の場合
     */
    public void forEach(BiConsumer<Pair.Side, ? super T> action) {
        Objects.requireNonNull(action, "action");
        a.ifPresent(value -> action.accept(Side.A, value));
        b.ifPresent(value -> action.accept(Side.B, value));
    }
    
    /**
     * 要素aと要素bを入れ替えた {@link Pair} オブジェクトを返します。<br>
     * 
     * @return 要素aと要素bを入れ替えた {@link Pair} オブジェクト
     */
    public Pair<T> reverse() {
        return Pair.flatOf(b, a);
    }
    
    @Override
    public boolean equals(Object o) {
        if (o instanceof Pair) {
            Pair<?> p = (Pair<?>) o;
            return Objects.equals(a, p.a) && Objects.equals(b, p.b);
        }
        return false;
    }
    
    @Override
    public int hashCode() {
        return Objects.hash(a, b);
    }
    
    @Override
    public String toString() {
        return String.format("(%s, %s)", a.orElse(null), b.orElse(null));
    }
}

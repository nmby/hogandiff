package xyz.hotchpotch.hogandiff;

import java.io.File;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.Arrays;
import java.util.Collections;
import java.util.HashMap;
import java.util.Map;
import java.util.Objects;
import java.util.Properties;
import java.util.function.Function;
import java.util.stream.Collectors;
import java.util.stream.Stream;

import org.apache.poi.ss.usermodel.IndexedColors;

/**
 * 現在のアプリケーション実行条件を表す不変クラスです。<br>
 * 
 * @author nmby
 * @since 0.1.0
 */
public class Context {
    
    // [static members] ********************************************************
    
    /**
     * コンテキストで管理するプロパティの種類を表す不変クラスです。<br>
     * このクラスは実質的に列挙ですが、総称型を利用するために通常のクラスとして実装されています。<br>
     * 
     * @param <T> プロパティ値の型
     * @author nmby
     * @since 0.1.0
     */
    public static class Props<T> {
        
        // [static members] ----------------------------------------------------
        
        /** 実行メニュー */
        public static final Props<Menu> CURR_MENU = new Props<>();
        
        /** 作業用フォルダのパス */
        public static final Props<Path> CURR_WORK_DIR = new Props<>();
        
        /** 比較対象ブック1 */
        public static final Props<File> CURR_FILE1 = new Props<>();
        
        /** 比較対象ブック2 */
        public static final Props<File> CURR_FILE2 = new Props<>();
        
        /** 比較対象シート1の名前 */
        public static final Props<String> CURR_SHEET_NAME1 = new Props<>();
        
        /** 比較対象シート2の名前 */
        public static final Props<String> CURR_SHEET_NAME2 = new Props<>();
        
        /** 比較において行の挿入／削除を考慮するか */
        public static final Props<Boolean> APP_CONSIDER_ROW_GAPS = new Props<>(
                true,
                "application.compare.considerRowGaps",
                true,
                Boolean::valueOf,
                String::valueOf);
        
        /** 比較において列の挿入／削除を考慮するか */
        public static final Props<Boolean> APP_CONSIDER_COLUMN_GAPS = new Props<>(
                true,
                "application.compare.considerColumnGaps",
                false,
                Boolean::valueOf,
                String::valueOf);
        
        /** 数式セルを数式ではなく値で比較するか */
        public static final Props<Boolean> APP_COMPARE_ON_VALUE = new Props<>(
                true,
                "application.compare.compareOnValue",
                false,
                Boolean::valueOf,
                String::valueOf);
        
        /** 比較結果のレポートとして、差分箇所に色を付けたシートを表示するか */
        public static final Props<Boolean> APP_SHOW_PAINTED_SHEETS = new Props<>(
                true,
                "application.report.showPaintedSheets",
                true,
                Boolean::valueOf,
                String::valueOf);
        
        /** 比較結果のレポートにおいて、余剰行／余剰列に付ける色 */
        public static final Props<Short> APP_REDUNDANT_COLOR = new Props<>(
                true,
                "application.report.redundantColor",
                IndexedColors.CORAL.getIndex(),
                Short::valueOf,
                String::valueOf);
        
        /** 比較結果のレポートにおいて、差分セルに付ける色 */
        public static final Props<Short> APP_DIFF_COLOR = new Props<>(
                true,
                "application.report.diffColor",
                IndexedColors.YELLOW.getIndex(),
                Short::valueOf,
                String::valueOf);
        
        /** 比較結果のレポートとして、差分内容を記したテキストを表示するか */
        public static final Props<Boolean> APP_SHOW_RESULT_TEXT = new Props<>(
                true,
                "application.report.showResultText",
                false,
                Boolean::valueOf,
                String::valueOf);
        
        /** 比較が完了したらこのアプリケーションを終了するか */
        public static final Props<Boolean> APP_EXIT_WHEN_FINISHED = new Props<>(
                true,
                "application.execution.exitWhenFinished",
                false,
                Boolean::valueOf,
                String::valueOf);
        
        /** 作業用フォルダの作成場所のパス */
        public static final Props<Path> SYS_WORK_DIR_BASE = new Props<>(
                true,
                "system.workDirBase",
                Paths.get(
                        System.getProperty("java.io.tmpdir"),
                        Props.class.getPackage().getName()),
                Paths::get,
                Path::toString);
        
        private static final Props<?>[] values;
        static {
            values = Arrays.stream(Props.class.getFields())
                    .filter(f -> f.getType() == Props.class)
                    .map(f -> {
                        try {
                            return f.get(null);
                        } catch (Exception e) {
                            return null;
                        }
                    })
                    .filter(v -> v != null)
                    .toArray(Props[]::new);
        }
        
        // [instance members] --------------------------------------------------
        
        /** プロパティファイルへの保存対象の場合は {@code true} */
        private final boolean storeOnFile;
        
        /** このプロパティのキー文字列 */
        private final String keyStr;
        
        /** このプロパティのデフォルト値 */
        private final T defaultValue;
        
        /** 文字列をこのプロパティの型に変換するデコーダ */
        private final Function<String, ? extends T> decoder;
        
        /** このプロパティの値を文字列型に変換するエンコーダ */
        private final Function<? super T, String> encoder;
        
        private Props() {
            this(false, null, null, null, null);
        }
        
        private Props(
                boolean storeOnFile,
                String keyStr,
                T defaultValue,
                Function<String, ? extends T> decoder,
                Function<? super T, String> encoder) {
            
            assert !storeOnFile || keyStr != null;
            assert !storeOnFile || decoder != null;
            assert !storeOnFile || encoder != null;
            
            this.storeOnFile = storeOnFile;
            this.keyStr = keyStr;
            this.defaultValue = defaultValue;
            this.decoder = decoder;
            this.encoder = encoder;
        }
    }
    
    /**
     * {@link Context} のビルダーです。<br>
     * 
     * @author nmby
     * @since 0.1.0
     */
    public static class Builder {
        
        // [static members] ----------------------------------------------------
        
        /**
         * プロパティセットをデフォルト値として用いてビルダーを作成します。<br>
         * 
         * @param defaults デフォルト値を保持するプロパティセット
         * @return ビルダー
         * @throws NullPointerException {@code defaults} が {@code null} の場合
         */
        public static Builder of(Properties defaults) {
            Objects.requireNonNull(defaults, "defaults");
            
            Map<Props<?>, Object> props = Stream.of(Props.values)
                    .filter(p -> p.storeOnFile)
                    .collect(Collectors.toMap(
                            Function.identity(),
                            p -> {
                                Object value = p.defaultValue;
                                if (defaults.containsKey(p.keyStr)) {
                                    String valueStr = defaults.getProperty(p.keyStr);
                                    try {
                                        value = p.decoder.apply(valueStr);
                                    } catch (RuntimeException e) {
                                        // nop : use p.defaultValue
                                    }
                                }
                                return value;
                            }));
            
            return new Builder(props);
        }
        
        // [instance members] --------------------------------------------------
        
        private final Map<Props<?>, Object> props;
        
        private Builder(Map<Props<?>, Object> props) {
            assert props != null;
            this.props = new HashMap<>(props);
        }
        
        /**
         * プロパティの値を指定します。<br>
         * 
         * @param <T> プロパティ値の型
         * @param key プロパティキー
         * @param value プロパティ値
         * @return このビルダー
         * @throws NullPointerException {@code key} が {@code null} の場合
         */
        public <T> Builder set(Props<T> key, T value) {
            Objects.requireNonNull(key, "key");
            props.put(key, value);
            return this;
        }
        
        /**
         * {@link Context} オブジェクトを作成します。<br>
         * 
         * @return {@link Context} オブジェクト
         * @throws IllegalStateException このビルダーの内容が不正な場合
         */
        public Context build() {
            if (props.get(Props.CURR_WORK_DIR) == null) {
                Path base = (Path) props.get(Props.SYS_WORK_DIR_BASE);
                String timestamp = LocalDateTime.now()
                        .format(DateTimeFormatter.ofPattern("yyyyMMdd-HHmmss-SSS"));
                props.put(Props.CURR_WORK_DIR, base.resolve(timestamp));
            }
            return new Context(this);
        }
    }
    
    // [instance members] ******************************************************
    
    private final Map<Props<?>, ?> props;
    
    private Context(Builder builder) {
        assert builder != null;
        this.props = Collections.unmodifiableMap(new HashMap<>(builder.props));
    }
    
    /**
     * 指定したプロパティの値を返します。<br>
     * 
     * @param <T> プロパティ値の型
     * @param key プロパティ
     * @return プロパティ値
     * @throws NullPointerException {@code key} が {@code null} の場合
     */
    @SuppressWarnings("unchecked")
    public <T> T get(Props<T> key) {
        Objects.requireNonNull(key, "key");
        return (T) props.get(key);
    }
    
    /**
     * この {@link Context} オブジェクトが表す実行条件のうちアプリケーションプロパティファイルに保存すべきプロパティ値を
     * {@link Properties} オブジェクトに抽出して返します。<br>
     * 
     * @return アプリケーションプロパティファイルに保存すべきプロパティ値を保持するプロパティセット
     */
    public Properties extractPropertiesToStore() {
        Properties properties = new Properties();
        for (Props<?> p : Props.values) {
            if (p.storeOnFile) {
                String valueStr = encode(p);
                properties.setProperty(p.keyStr, valueStr);
            }
        }
        return properties;
    }
    
    private <U> String encode(Props<U> props) {
        return props.encoder.apply(get(props));
    }
}

package xyz.hotchpotch.hogandiff.excel.xssf;

import java.io.InputStream;
import java.nio.file.FileSystem;
import java.nio.file.FileSystems;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.Collections;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.NoSuchElementException;
import java.util.Objects;
import java.util.Optional;
import java.util.stream.Collectors;

import org.xml.sax.Attributes;
import org.xml.sax.InputSource;
import org.xml.sax.XMLReader;
import org.xml.sax.helpers.DefaultHandler;
import org.xml.sax.helpers.XMLReaderFactory;

import xyz.hotchpotch.hogandiff.ApplicationException;

/**
 * XSSF（.xlsx/.xlsm）形式のExcelブックのシート情報を保持する不変クラスです。<br>
 * 
 * @author nmby
 * @since 0.3.2
 */
public class XSSFSheetEntryManager {
    
    // [static members] ********************************************************
    
    /**
     * zipファイルとしての.xlsx/.xlsmファイルから次のエントリを読み込み、
     * シート名の一覧、および、シート名とシートId（relId）のマップを抽出します。<br>
     * <pre>
     * *.xlsx
     *   +-xl
     *     +-workbook.xml
     * </pre>
     * 
     * @author nmby
     * @since 0.3.2
     */
    private static class Handler1 extends DefaultHandler {
        
        // [static members] ----------------------------------------------------
        
        // [instance members] --------------------------------------------------
        
        private List<String> names;
        private Map<String, String> nameToId;
        
        @Override
        public void startDocument() {
            names = new ArrayList<>();
            nameToId = new HashMap<>();
        }
        
        @Override
        public void startElement(String uri, String localName, String qName, Attributes attributes) {
            if ("sheet".equals(qName)) {
                names.add(attributes.getValue("name"));
                nameToId.put(
                        attributes.getValue("name"),
                        attributes.getValue("r:id"));
            }
        }
    }
    
    /**
     * zipファイルとしての.xlsx/.xlsmファイルから次のエントリを読み込み、
     * シート名とソースエントリのパスのマップを抽出します。<br>
     * <pre>
     * *.xlsx
     *   +-xl
     *     +-_rels
     *       +-workbook.xml.rels
     * </pre>
     * 
     * @author nmby
     * @since 0.3.2
     */
    private static class Handler2 extends DefaultHandler {
        
        // [static members] ----------------------------------------------------
        
        // [instance members] --------------------------------------------------
        
        private Map<String, String> idToSource;
        
        @Override
        public void startDocument() {
            idToSource = new HashMap<>();
        }
        
        @Override
        public void startElement(String uri, String localName, String qName, Attributes attributes) {
            if ("Relationship".equals(qName)) {
                idToSource.put(
                        attributes.getValue("Id"),
                        "xl/" + attributes.getValue("Target"));
            }
        }
    }
    
    /**
     * 指定された .xlsx/.xlsm ファイルを読み込んでシート情報を抽出し、
     * 抽出したシート情報を保持する {@link XSSFSheetEntryManager} オブジェクトを生成して返します。<br>
     * 
     * @param target 対象のExcelファイル（.xlsx/.xlsm 形式）
     * @return 対象Excelブックのシート情報を保持する新しい {@link XSSFSheetEntryManager} オブジェクト
     * @throws ApplicationException 処理に失敗した場合
     * @throws NullPointerException {@code target} が {@code null} の場合
     */
    public static XSSFSheetEntryManager generate(Path target) throws ApplicationException {
        Objects.requireNonNull(target, "target");
        
        try (FileSystem fs = FileSystems.newFileSystem(target, null)) {
            Handler1 handler1 = new Handler1();
            Handler2 handler2 = new Handler2();
            
            try (InputStream is = Files.newInputStream(fs.getPath("xl/workbook.xml"))) {
                InputSource source = new InputSource(is);
                XMLReader parser = XMLReaderFactory.createXMLReader();
                parser.setContentHandler(handler1);
                parser.parse(source);
            }
            
            try (InputStream is = Files.newInputStream(fs.getPath("xl/_rels/workbook.xml.rels"))) {
                InputSource source = new InputSource(is);
                XMLReader parser = XMLReaderFactory.createXMLReader();
                parser.setContentHandler(handler2);
                parser.parse(source);
            }
            
            return new XSSFSheetEntryManager(
                    handler1.names,
                    handler1.nameToId,
                    handler2.idToSource);
            
        } catch (Exception e) {
            throw new ApplicationException("Excelブックの解析に失敗しました。\n" + target.toString());
        }
    }
    
    // [instance members] ******************************************************
    
    private final List<String> names;
    private final Map<String, String> nameToId;
    private final Map<String, String> idToSource;
    
    private XSSFSheetEntryManager(
            List<String> names,
            Map<String, String> nameToId,
            Map<String, String> idToSource) {
        
        assert names != null;
        assert nameToId != null;
        assert idToSource != null;
        assert names.stream().allMatch(nameToId::containsKey);
        assert names.stream().map(nameToId::get).allMatch(idToSource::containsKey);
        
        // このオブジェクトを不変にするために防御的コピーしたうえで変更不可コレクションでラップする。
        this.names = Collections.unmodifiableList(new ArrayList<>(names));
        this.nameToId = Collections.unmodifiableMap(new HashMap<>(nameToId));
        this.idToSource = Collections.unmodifiableMap(new HashMap<>(idToSource));
    }
    
    /**
     * シート名のリストを返します。
     * 返されるリストには、ワークシートおよびグラフシートの双方の名前が含まれます。<br>
     * 
     * @return ワークシート名およびグラフシート名のリスト
     */
    public List<String> getNames() {
        return names;
    }
    
    /**
     * ワークシート名のリストを返します。
     * 返されるリストにグラフシートの名前は含まれません。<br>
     * 
     * @return ワークシート名のリスト
     */
    public List<String> getWorksheetNames() {
        return names.stream()
                .filter(name -> getSourceByName(name).startsWith("xl/worksheets/"))
                .collect(Collectors.toList());
    }
    
    /**
     * シート名に対応するシートIdを返します。<br>
     * 
     * @param name シート名
     * @return シート名に対応するシートId
     * @throws NullPointerException {@code name} が {@code null} の場合
     * @throws NoSuchElementException {@code name} に該当するシートが存在しない場合
     */
    public String getIdByName(String name) {
        Objects.requireNonNull(name, "name");
        return Optional.ofNullable(nameToId.get(name))
                .orElseThrow(() -> new NoSuchElementException("name: " + name));
    }
    
    /**
     * シート名に対応するソースエントリを返します。<br>
     * 
     * @param name シート名
     * @return シート名に対応するソースエントリ
     * @throws NullPointerException {@code name} が {@code null} の場合
     * @throws NoSuchElementException {@code name} に該当するシートが存在しない場合
     */
    public String getSourceByName(String name) {
        String id = getIdByName(name);
        return Optional.ofNullable(idToSource.get(id))
                .orElseThrow(() -> new AssertionError("id: " + id));
    }
}

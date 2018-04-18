package xyz.hotchpotch.hogandiff.poi;

import java.io.File;
import java.io.FileInputStream;
import java.util.EnumSet;
import java.util.HashSet;
import java.util.List;
import java.util.NoSuchElementException;
import java.util.Objects;
import java.util.Set;
import java.util.stream.Collectors;
import java.util.stream.IntStream;

import org.apache.poi.hssf.eventusermodel.HSSFEventFactory;
import org.apache.poi.hssf.eventusermodel.HSSFListener;
import org.apache.poi.hssf.eventusermodel.HSSFRequest;
import org.apache.poi.hssf.record.BOFRecord;
import org.apache.poi.hssf.record.BoolErrRecord;
import org.apache.poi.hssf.record.BoundSheetRecord;
import org.apache.poi.hssf.record.CellRecord;
import org.apache.poi.hssf.record.EOFRecord;
import org.apache.poi.hssf.record.FormulaRecord;
import org.apache.poi.hssf.record.LabelSSTRecord;
import org.apache.poi.hssf.record.NumberRecord;
import org.apache.poi.hssf.record.RKRecord;
import org.apache.poi.hssf.record.Record;
import org.apache.poi.hssf.record.SSTRecord;
import org.apache.poi.hssf.record.StringRecord;
import org.apache.poi.hssf.record.common.UnicodeString;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.formula.eval.ErrorEval;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.FormulaError;
import org.apache.poi.ss.util.NumberToTextConverter;

import xyz.hotchpotch.hogandiff.ApplicationException;
import xyz.hotchpotch.hogandiff.common.BookType;
import xyz.hotchpotch.hogandiff.common.CellReplica;

/**
 * HSSF（.xls）形式のExcelブックからPOIのイベントモデルAPIを使用してシートデータを読み込むための
 * {@link SheetLoader} の実装です。<br>
 * 
 * @author nmby
 * @since 0.1.0
 */
/*package*/ class HSSFSheetLoaderWithEventApi implements SheetLoader {
    
    // [static members] ********************************************************
    
    private static enum ProcessingPhase {
        
        // [static members] ----------------------------------------------------
        
        SEARCHING_SHEET,
        READING_SST_DATA,
        WAITING_SHEET_BODY,
        READING_SHEET_DATA,
        COMPLETED;
        
        // [instance members] --------------------------------------------------
    }
    
    private static class HSSFSheetLoadingListener implements HSSFListener {
        
        // [static members] ----------------------------------------------------
        
        // [instance members] --------------------------------------------------
        
        private final String sheetName;
        private final boolean extractCachedValue;
        
        private ProcessingPhase phase = ProcessingPhase.SEARCHING_SHEET;
        private int sheetIdx;
        private int currIdx;
        private FormulaRecord prevFormula;
        private List<String> sst;
        private Set<CellReplica> cells = new HashSet<>();
        
        private HSSFSheetLoadingListener(String sheetName, boolean extractCachedValue) {
            assert sheetName != null;
            this.sheetName = sheetName;
            this.extractCachedValue = extractCachedValue;
        }
        
        // 我ながらこの醜いコードはもうちょっとどうにかならないんだろうか・・・ orz
        @Override
        public void processRecord(Record record) {
            switch (phase) {
            case SEARCHING_SHEET:
                if (record.getSid() == BoundSheetRecord.sid) {
                    BoundSheetRecord boundSheetRec = (BoundSheetRecord) record;
                    if (sheetName.equals(boundSheetRec.getSheetname())) {
                        phase = ProcessingPhase.READING_SST_DATA;
                    } else {
                        sheetIdx++;
                    }
                } else if (record.getSid() == EOFRecord.sid) {
                    throw new NoSuchElementException(sheetName);
                }
                break;
            
            case READING_SST_DATA:
                if (record.getSid() == SSTRecord.sid) {
                    SSTRecord sstRec = (SSTRecord) record;
                    sst = IntStream.range(0, sstRec.getNumUniqueStrings())
                            .mapToObj(sstRec::getString)
                            .map(UnicodeString::getString)
                            .collect(Collectors.toList());
                    phase = ProcessingPhase.WAITING_SHEET_BODY;
                } else if (record.getSid() == EOFRecord.sid) {
                    throw new AssertionError("no sst record.");
                }
                break;
            
            case WAITING_SHEET_BODY:
                if (record.getSid() == BOFRecord.sid) {
                    BOFRecord bofRec = (BOFRecord) record;
                    if (bofRec.getType() == BOFRecord.TYPE_WORKSHEET
                            || bofRec.getType() == BOFRecord.TYPE_CHART) {
                        
                        if (currIdx == sheetIdx && bofRec.getType() == BOFRecord.TYPE_WORKSHEET) {
                            phase = ProcessingPhase.READING_SHEET_DATA;
                        } else if (currIdx == sheetIdx && bofRec.getType() == BOFRecord.TYPE_CHART) {
                            // グラフシートの場合は空のセルデータセットを返すこととする。
                            phase = ProcessingPhase.COMPLETED;
                        } else if (currIdx < sheetIdx) {
                            currIdx++;
                        } else {
                            throw new AssertionError("no sheet body.");
                        }
                    }
                }
                break;
            
            case READING_SHEET_DATA:
                if (record instanceof CellRecord) {
                    CellRecord cellRec = (CellRecord) record;
                    String value = null;
                    
                    switch (cellRec.getSid()) {
                    case LabelSSTRecord.sid:
                        LabelSSTRecord labelSSTRec = (LabelSSTRecord) cellRec;
                        value = sst.get(labelSSTRec.getSSTIndex());
                        break;
                    
                    case NumberRecord.sid:
                        NumberRecord numberRec = (NumberRecord) cellRec;
                        value = NumberToTextConverter.toText(numberRec.getValue());
                        break;
                    
                    case RKRecord.sid:
                        RKRecord rKRec = (RKRecord) cellRec;
                        value = NumberToTextConverter.toText(rKRec.getRKNumber());
                        break;
                    
                    case BoolErrRecord.sid:
                        BoolErrRecord boolErrRec = (BoolErrRecord) cellRec;
                        if (boolErrRec.isError()) {
                            value = FormulaError.forInt(boolErrRec.getErrorValue()).getString();
                        } else {
                            value = String.valueOf(boolErrRec.getBooleanValue());
                        }
                        break;
                    
                    case FormulaRecord.sid:
                        FormulaRecord formulaRec = (FormulaRecord) cellRec;
                        if (extractCachedValue) {
                            // 汚いコードだけど許して... （もっと良い方法教えてほしい...）
                            @SuppressWarnings("deprecation")
                            CellType type = CellType.forInt(formulaRec.getCachedResultType());
                            switch (type) {
                            // see org.apache.poi.hssf.record.FormulaRecord.SpecialCachedValue#getValueType()
                            case NUMERIC:
                                value = NumberToTextConverter.toText(formulaRec.getValue());
                                break;
                            case STRING:
                                assert formulaRec.hasCachedResultString();
                                prevFormula = formulaRec;
                                break;
                            case BOOLEAN:
                                value = String.valueOf(formulaRec.getCachedBooleanValue());
                                break;
                            case ERROR:
                                value = ErrorEval.getText(formulaRec.getCachedErrorValue());
                                break;
                            default:
                                throw new AssertionError(type);
                            }
                        } else {
                            value = getFormulaString(formulaRec);
                        }
                        break;
                    
                    default:
                        throw new AssertionError(cellRec.getSid());
                    }
                    
                    if (prevFormula == null) {
                        assert value != null;
                        cells.add(CellReplica.of(
                                cellRec.getRow(),
                                cellRec.getColumn(),
                                value));
                    }
                    
                } else if (record instanceof StringRecord && prevFormula != null) {
                    StringRecord stringRec = (StringRecord) record;
                    cells.add(CellReplica.of(
                            prevFormula.getRow(),
                            prevFormula.getColumn(),
                            stringRec.getString()));
                    prevFormula = null;
                    
                } else if (record instanceof EOFRecord) {
                    phase = ProcessingPhase.COMPLETED;
                }
                break;
            
            case COMPLETED:
                // nop
                break;
            
            default:
                throw new AssertionError(phase);
            }
        }
    }
    
    private static String getFormulaString(FormulaRecord rec) {
        // 次のサイトに貴重な参考情報があるものの、数式文字列を得るためにユーザーモデルAPI（HSSFWorkbook）を
        // 利用する必要があり、わざわざイベントモデルAPIで頑張っている意味がなくなってしまう。
        // http://www.ne.jp/asahi/hishidama/home/tech/apache/poi/cell.html#h_toFormulaString
        // そのうち頑張ってコーディングすることとする。。。
        // TODO: coding
        return "[formula]";
    }
    
    /** このクラスがサポートするブック形式 */
    private static final Set<BookType> supported = EnumSet.of(BookType.XLS);
    
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
    public static HSSFSheetLoaderWithEventApi of(boolean extractCachedValue) {
        return new HSSFSheetLoaderWithEventApi(extractCachedValue);
    }
    
    // [instance members] ******************************************************
    
    private final boolean extractCachedValue;
    
    private HSSFSheetLoaderWithEventApi(boolean extractCachedValue) {
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
        
        try (FileInputStream fin = new FileInputStream(book);
                POIFSFileSystem poifs = new POIFSFileSystem(fin)) {
            
            HSSFRequest req = new HSSFRequest();
            HSSFSheetLoadingListener listener = new HSSFSheetLoadingListener(sheetName, extractCachedValue);
            req.addListenerForAllRecords(listener);
            HSSFEventFactory factory = new HSSFEventFactory();
            factory.abortableProcessWorkbookEvents(req, poifs);
            return listener.cells;
            
        } catch (NoSuchElementException e) {
            throw e;
        } catch (Exception e) {
            throw new ApplicationException(e);
        }
    }
}

package xyz.hotchpotch.hogandiff.excel.xssf.readers;

import javax.xml.stream.XMLEventReader;
import javax.xml.stream.XMLStreamException;
import javax.xml.stream.events.XMLEvent;

/**
 * {@link XMLEventReader} の次のオペレーションについて1回目の実行結果をキャッシュする基底実装です。<br>
 * <ul>
 *   <li>{@link XMLEventReader#hasNext()}</li>
 *   <li>{@link XMLEventReader#peek()}</li>
 * </ul>
 * 
 * @author nmby
 * @since 0.4.0
 */
/*package*/ abstract class AbstractCachingReader implements XMLEventReader {
    
    // [static members] ********************************************************
    
    // [instance members] ******************************************************
    
    private Boolean cachedHasNext = null;
    private XMLEvent cachedPeek = null;
    
    protected AbstractCachingReader() {
    }
    
    /**
     * ある状態において {@link #hasNext()} が始めて実行されるときに呼び出され、結果がキャッシュされます。<br>
     * 同じ状態において {@link #hasNext()} が再度実行された場合はキャッシュされた値が返され、
     * このオペレーションは呼び出されません。<br>
     * 
     * @return 次の要素が存在する場合は {@code true}
     */
    protected abstract boolean hasNext2();
    
    /**
     * ある状態において {@link #peek()} が始めて実行されるときに呼び出され、結果がキャッシュされます。<br>
     * 同じ状態において {@link #peek()} が再度実行された場合はキャッシュされた値が返され、
     * このオペレーションは呼び出されません。<br>
     * 
     * @return 次の要素（ただしソースから削除しません）。次の要素が存在しない場合は {@code null}
     * @throws XMLStreamException XMLイベントの解析に失敗した場合
     */
    protected abstract XMLEvent peek2() throws XMLStreamException;
    
    /**
     * {@link #nextEvent()} が実行されるときに呼び出されます。<br>
     * 
     * @return 次の要素（ソースから削除します）。次の要素が存在しない場合は {@code null}
     * @throws XMLStreamException XMLイベントの解析に失敗した場合
     */
    protected abstract XMLEvent nextEvent2() throws XMLStreamException;
    
    @Override
    public boolean hasNext() {
        if (cachedHasNext == null) {
            cachedHasNext = hasNext2();
        }
        return cachedHasNext;
    }
    
    @Override
    public XMLEvent peek() throws XMLStreamException {
        if (!hasNext()) {
            return null;
        }
        if (cachedPeek == null) {
            cachedPeek = peek2();
        }
        return cachedPeek;
    }
    
    @Override
    public XMLEvent nextEvent() throws XMLStreamException {
        if (!hasNext()) {
            return null;
        }
        XMLEvent event = nextEvent2();
        cachedHasNext = null;
        cachedPeek = null;
        return event;
    }
    
    @Override
    public Object next() {
        try {
            return nextEvent();
        } catch (XMLStreamException e) {
            throw new RuntimeException(new XMLStreamException());
        }
    }
    
    /**
     * このオペレーションはサポートされません。<br>
     * 
     * @throws UnsupportedOperationException このオペレーションを実行した場合
     */
    @Override
    public String getElementText() throws XMLStreamException {
        throw new UnsupportedOperationException();
    }
    
    /**
     * このオペレーションはサポートされません。<br>
     * 
     * @throws UnsupportedOperationException このオペレーションを実行した場合
     */
    @Override
    public XMLEvent nextTag() throws XMLStreamException {
        throw new UnsupportedOperationException();
    }
    
    /**
     * このオペレーションはサポートされません。<br>
     * 
     * @throws UnsupportedOperationException このオペレーションを実行した場合
     */
    @Override
    public Object getProperty(String name) throws IllegalArgumentException {
        throw new UnsupportedOperationException();
    }
}

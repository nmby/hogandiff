package xyz.hotchpotch.hogandiff.excel.xssf.readers;

import java.util.ArrayDeque;
import java.util.ArrayList;
import java.util.Deque;
import java.util.Iterator;
import java.util.List;
import java.util.Objects;
import java.util.function.BiPredicate;
import java.util.function.Predicate;

import javax.xml.namespace.QName;
import javax.xml.stream.XMLEventReader;
import javax.xml.stream.XMLStreamException;
import javax.xml.stream.events.StartElement;
import javax.xml.stream.events.XMLEvent;

/**
 * 指定された要素を読み飛ばす {@link XMLEventReader} の実装です。<br>
 * 
 * @author nmby
 * @since 0.4.0
 */
public class FilteringReader extends AbstractCachingReader {
    
    // [static members] ********************************************************
    
    /**
     * {@link FilteringReader} のビルダーです。<br>
     * 
     * @author nmby
     * @since 0.4.0
     */
    public static class Builder {
        
        // [static members] ----------------------------------------------------
        
        // [instance members] --------------------------------------------------
        
        private final XMLEventReader source;
        private final List<BiPredicate<? super Deque<? super QName>, ? super StartElement>> filters = new ArrayList<>();
        
        private Builder(XMLEventReader source) {
            assert source != null;
            this.source = source;
        }
        
        /**
         * このビルダーにフィルタを追加します。<br>
         * 例えば div 要素の配下の table 要素の配下の font 要素を読み飛ばしたい場合は
         * {div, table, font} を指定します。<br>
         * 
         * @param qNames 読み飛ばす要素を表すQName階層
         * @return このビルダー
         * @throws NullPointerException {@code qNames} が {@code null} の場合
         */
        public Builder addFilter(QName... qNames) {
            Objects.requireNonNull(qNames, "qNames");
            if (qNames.length == 0) {
                throw new IllegalArgumentException();
            }
            
            filters.add((currTree, start) -> {
                int i = qNames.length - 1;
                if (!qNames[i].equals(start.getName())) {
                    return false;
                }
                
                if (currTree.size() + 1 < qNames.length) {
                    return false;
                }
                
                Iterator<? super QName> itr = currTree.descendingIterator();
                while (0 < i) {
                    i--;
                    if (!qNames[i].equals(itr.next())) {
                        return false;
                    }
                }
                return true;
            });
            
            return this;
        }
        
        /**
         * このビルダーにフィルタを追加します。<br>
         * 
         * @param filter 除外する要素の場合に {@code true} を返す関数
         * @return このビルダー
         * @throws NullPointerException {@code filter} が {@code null} の場合
         */
        public Builder addFilter(Predicate<? super StartElement> filter) {
            Objects.requireNonNull(filter, "filter");
            
            filters.add((currTree, start) -> filter.test(start));
            
            return this;
        }
        
        /**
         * このビルダーにフィルタを追加します。<br>
         * 
         * @param filter 現在の要素ツリーと次の開始要素を受け取り除外対象の場合に {@code true} を返す関数
         * @return このビルダー
         * @throws NullPointerException {@code filter} が {@code null} の場合
         */
        public Builder addFilter(
                BiPredicate<? super Deque<? super QName>, ? super StartElement> filter) {
            
            Objects.requireNonNull(filter, "filter");
            
            filters.add((currTree, start) -> filter.test(currTree, start));
            
            return this;
        }
        
        /**
         * このビルダーから {@link FilteringReader} を構成します。<br>
         * 
         * @return 新しいリーダー
         */
        public XMLEventReader build() {
            return new FilteringReader(this);
        }
    }
    
    /**
     * ビルダーを生成します。<br>
     * 
     * @param source ソースリーダー
     * @return 新しいビルダー
     * @throws NullPointerException {@code source} が {@code null} の場合
     */
    public static Builder builder(XMLEventReader source) {
        Objects.requireNonNull(source, "source");
        return new Builder(source);
    }
    
    // [instance members] ******************************************************
    
    private final XMLEventReader source;
    private final List<BiPredicate<? super Deque<? super QName>, ? super StartElement>> filters;
    private final Deque<QName> currTree = new ArrayDeque<>();
    
    private FilteringReader(Builder builder) {
        super();
        
        assert builder != null;
        
        this.source = builder.source;
        this.filters = builder.filters;
    }
    
    private void seekNext() throws XMLStreamException {
        while (source.hasNext() && source.peek().isStartElement()) {
            StartElement next = source.peek().asStartElement();
            if (filters.stream().noneMatch(filter -> filter.test(currTree, next))) {
                return;
            }
            int depth = 0;
            do {
                XMLEvent event = source.nextEvent();
                if (event.isStartElement()) {
                    depth++;
                } else if (event.isEndElement()) {
                    depth--;
                }
            } while (source.hasNext() && 0 < depth);
            if (depth != 0) {
                throw new XMLStreamException();
            }
        }
    }
    
    @Override
    protected boolean hasNext2() {
        try {
            seekNext();
        } catch (XMLStreamException e) {
            throw new RuntimeException(e);
        }
        return source.hasNext();
    }
    
    @Override
    protected XMLEvent peek2() throws XMLStreamException {
        return source.peek();
    }
    
    @Override
    protected XMLEvent nextEvent2() throws XMLStreamException {
        XMLEvent event = source.nextEvent();
        if (event.isStartElement()) {
            currTree.addLast(event.asStartElement().getName());
        } else if (event.isEndElement()) {
            currTree.removeLast();
        }
        return event;
    }
    
    @Override
    public void close() throws XMLStreamException {
        source.close();
    }
}

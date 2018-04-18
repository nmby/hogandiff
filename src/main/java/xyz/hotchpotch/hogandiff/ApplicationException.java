package xyz.hotchpotch.hogandiff;

/**
 * アプリケーションの処理に失敗したことを表すチェック例外です。<br>
 * 
 * @author nmby
 * @since 0.1.0
 */
public class ApplicationException extends Exception {
    
    // [static members] ********************************************************
    
    // [instance members] ******************************************************
    
    /**
     * 例外オブジェクトを生成します。<br>
     */
    public ApplicationException() {
        super();
    }
    
    /**
     * 例外オブジェクトを生成します。<br>
     * 
     * @param msg 例外メッセージ
     */
    public ApplicationException(String msg) {
        super(msg);
    }
    
    /**
     * 例外オブジェクトを生成します。<br>
     * 
     * @param cause 例外の原因
     */
    public ApplicationException(Throwable cause) {
        super(cause);
    }
    
    /**
     * 例外オブジェクトを生成します。<br>
     * 
     * @param msg 例外メッセージ
     * @param cause 例外の原因
     */
    public ApplicationException(String msg, Throwable cause) {
        super(msg, cause);
    }
}

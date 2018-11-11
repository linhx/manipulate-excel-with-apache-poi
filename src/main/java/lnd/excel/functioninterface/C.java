package lnd.excel.functioninterface;

/**
 * @author linhnguyendinh
 */
@FunctionalInterface
public interface C<T> {
    void accept(T t) throws Exception;
}

package lnd.excel.functioninterface;

/**
 * @author linhnguyendinh
 */
@FunctionalInterface
public interface BiC<T, U> {
    void accept(T t, U u) throws Exception;
}

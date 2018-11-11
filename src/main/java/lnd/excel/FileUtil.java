package lnd.excel;

import lnd.excel.functioninterface.BiC;
import org.apache.commons.collections4.CollectionUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;

import java.util.List;

/**
 * @author linhnguyendinh
 */
public class FileUtil {

    public static final String EXCEL_NAME_REGEX_INVALID_CHAR = "[:\\\\/?*\\[\\]]";

    /**
     * get Cell by row index and column index
     *
     * @param sheet the sheet
     * @param rowIndex the row index
     * @param columnIndex the column index
     * @return the cell
     */
    public static Cell cell(Sheet sheet, int rowIndex, int columnIndex) {
        Row row = sheet.getRow(rowIndex);
        if (row == null) {
            row = sheet.createRow(rowIndex);
        }
        Cell cell = row.getCell(columnIndex);
        if (cell == null) {
            cell = row.createCell(columnIndex);
        }
        return cell;
    }

    /**
     * get cell by name defined in excel file
     *
     * @param sheet the worksheet
     * @param name the named range (refer {@link Name})
     * @param offsetY the row offset
     * @param offsetX the column offset
     * @return the cell
     */
    public static Cell cell(Sheet sheet, String name, int offsetY, int offsetX) {
        Name n = sheet.getWorkbook().getName(name);
        CellReference cellReference = new CellReference(n == null? name: n.getRefersToFormula());
        return FileUtil.cell(sheet, cellReference.getRow() + offsetY, cellReference.getCol() + offsetX);
    }

    /**
     * get cell by name defined in excel file
     *
     * @param sheet the worksheet
     * @param name the named range (refer {@link Name})
     * @return the cell
     */
    public static Cell cell(Sheet sheet, String name) {
        return cell(sheet, name, 0, 0);
    }

    /**
     * copy style, content from source cell to destination cell
     * @param srcCell the source cell
     * @param destCell the destination cell
     */
    private static void copyCell(Cell srcCell, Cell destCell) {
        CellType cellType = srcCell.getCellTypeEnum();

        switch (cellType) {
            case STRING:
                destCell.setCellValue(srcCell.getRichStringCellValue());
                break;
            case BOOLEAN:
                destCell.setCellValue(srcCell.getBooleanCellValue());
                break;
            case NUMERIC:
                destCell.setCellValue(srcCell.getNumericCellValue());
                break;
            case FORMULA:
                destCell.setCellValue(srcCell.getRichStringCellValue());
                break;
            case ERROR:
                destCell.setCellValue(srcCell.getCellFormula());
                break;
            default:
                break;
        }

        destCell.setCellStyle(srcCell.getCellStyle());
        if(srcCell.getCellComment() != null) destCell.setCellComment(srcCell.getCellComment());
        if(srcCell.getHyperlink() != null) destCell.setHyperlink(srcCell.getHyperlink());
    }

    /**
     * copy sheet setup (PageSetup, Header, Footer)
     * @param srcSheet source sheet
     * @param destSheet destination sheet
     */
    private static void copySheetSetup(Sheet srcSheet, Sheet destSheet) {
        PrintSetup printSetup = srcSheet.getPrintSetup();
        Header header = srcSheet.getHeader();
        Footer footer = srcSheet.getFooter();

        destSheet.setAutobreaks(srcSheet.getAutobreaks());
        destSheet.setMargin(Sheet.LeftMargin, srcSheet.getMargin(Sheet.LeftMargin));
        destSheet.setMargin(Sheet.RightMargin, srcSheet.getMargin(Sheet.RightMargin));
        destSheet.setMargin(Sheet.TopMargin, srcSheet.getMargin(Sheet.TopMargin));
        destSheet.setMargin(Sheet.BottomMargin, srcSheet.getMargin(Sheet.BottomMargin));
        destSheet.setMargin(Sheet.HeaderMargin, srcSheet.getMargin(Sheet.HeaderMargin));
        destSheet.setMargin(Sheet.FooterMargin, srcSheet.getMargin(Sheet.FooterMargin));

        destSheet.setHorizontallyCenter(srcSheet.getHorizontallyCenter());
        destSheet.setVerticallyCenter(srcSheet.getVerticallyCenter());
        destSheet.setRepeatingColumns(srcSheet.getRepeatingColumns());
        destSheet.setRepeatingRows(srcSheet.getRepeatingRows());

        PrintSetup clonePrintSetup = destSheet.getPrintSetup();
        clonePrintSetup.setCopies(printSetup.getCopies());
        clonePrintSetup.setDraft(printSetup.getDraft());
        clonePrintSetup.setFitHeight(printSetup.getFitHeight());
        clonePrintSetup.setFitWidth(printSetup.getFitWidth());
        clonePrintSetup.setFooterMargin(printSetup.getFooterMargin());
        clonePrintSetup.setHeaderMargin(printSetup.getHeaderMargin());
        clonePrintSetup.setHResolution(printSetup.getHResolution());
        clonePrintSetup.setVResolution(printSetup.getVResolution());
        clonePrintSetup.setLandscape(printSetup.getLandscape());
        clonePrintSetup.setLeftToRight(printSetup.getLeftToRight());
        clonePrintSetup.setNoColor(printSetup.getNoColor());
        clonePrintSetup.setNoOrientation(printSetup.getNoOrientation());
        clonePrintSetup.setNotes(printSetup.getNotes());
        clonePrintSetup.setPageStart(printSetup.getPageStart());
        clonePrintSetup.setPaperSize(printSetup.getPaperSize());
        clonePrintSetup.setScale(printSetup.getScale());
        clonePrintSetup.setUsePage(printSetup.getUsePage());
        clonePrintSetup.setValidSettings(printSetup.getValidSettings());

        Header cloneHeader = destSheet.getHeader();
        cloneHeader.setCenter(header.getCenter());
        cloneHeader.setLeft(header.getLeft());
        cloneHeader.setRight(header.getRight());

        Footer cloneFooter = destSheet.getFooter();
        cloneFooter.setCenter(footer.getCenter());
        cloneFooter.setLeft(footer.getLeft());
        cloneFooter.setRight(footer.getRight());
    }

    /**
     * copy sheet and handle printer
     *
     * @param sheet the sheet template
     * @param consumer the print handler
     * @param datas list of data, each element print on a sheet
     * @param <T> the type of the datas
     */
    public static <T> void copySheet(Sheet sheet, BiC<Sheet, T> consumer, List<T> datas) throws Exception {
        if (CollectionUtils.isEmpty(datas)) return;
        Workbook workbook = sheet.getWorkbook();
        int sheetIndex = workbook.getSheetIndex(sheet);

        for (T data: datas) {
            Sheet sheetClone = workbook.cloneSheet(sheetIndex);
            // copy sheet setup
            FileUtil.copySheetSetup(sheet, sheetClone);
            // handle printer
            consumer.accept(sheetClone, data);
        }
        //remove sheet template
        workbook.removeSheetAt(sheetIndex);
    }

    /**
     * Copy and paste a range down to an interval addOffsetY
     *
     * @param sheet the worksheet
     * @param name the named range (refer {@link Name})
     * @param addOffsetY the add offset
     * @param consumer the new range handler
     * @param datas the data for new range handler
     * @param <T> the data type of the 'datas'
     */
    public static <T> void verticalCopyRange(Sheet sheet, String name, int addOffsetY, BiC<Range, T> consumer, List<T> datas) throws Exception {
        FileUtil.verticalCopyRange(sheet, name, addOffsetY, consumer, datas, false);
    }

    /**
     * Copy and insert a range down to an interval addOffsetY
     *
     * @param sheet the worksheet
     * @param name the named range (refer {@link Name})
     * @param addOffsetY the add offset
     * @param consumer the new range handler
     * @param datas the data for new range handler
     * @param <T> the data type of the 'datas'
     */
    public static <T> void verticalCopyInsertRange(Sheet sheet, String name, int addOffsetY, BiC<Range, T> consumer, List<T> datas) throws Exception {
        FileUtil.verticalCopyRange(sheet, name, addOffsetY, consumer, datas, true);
    }

    /**
     * Copy a range down to an interval addOffsetY
     *
     * @param sheet the worksheet
     * @param name the named range (refer {@link Name})
     * @param addOffsetY the add offset
     * @param consumer the new range handler
     * @param datas the data for new range handler
     * @param <T> the data type of the 'datas'
     * @param copyInsert is copy then insert below. in case there're existed rows below, avoid override below content
     */
    public static <T> void verticalCopyRange(Sheet sheet, String name, int addOffsetY, BiC<Range, T> consumer, List<T> datas, boolean copyInsert) throws Exception {
        if (CollectionUtils.isEmpty(datas)) return;

        T firstData = datas.remove(0);
        Range originalRange = new Range(sheet, 0, 0, name);
        int addOffset = addOffsetY;
        for (T data: datas) {
            Range rangeClone = copyInsert? originalRange.verticalCopyInsert(addOffset): originalRange.verticalCopy(addOffset);
            // handle printer
            consumer.accept(rangeClone, data);
            addOffset = rangeClone.getShiftY();
        }
        // print template sheet at last, for reason keep format
        consumer.accept(originalRange, firstData);
    }

    /**
     * Copy and paste a range to an interval addOffsetX
     *
     * @param sheet the worksheet
     * @param name the named range (refer {@link Name})
     * @param addOffsetX the add offset
     * @param consumer the new range handler
     * @param datas the data for new range handler
     * @param <T> the data type of the 'datas'
     */
    public static <T> void horizontalCopyRange(Sheet sheet, String name, int addOffsetX, BiC<Range, T> consumer, List<T> datas) throws Exception {
        if (CollectionUtils.isEmpty(datas)) return;

        T firstData = datas.remove(0);
        Range originalRange = new Range(sheet, 0, 0, name);
        int addOffset = addOffsetX;
        for (T data: datas) {
            Range rangeClone = originalRange.horizontalCopy(addOffset);
            // handle printer
            consumer.accept(rangeClone, data);
            addOffset = rangeClone.getShiftX();
        }
        // print template sheet at last, for reason keep format
        consumer.accept(originalRange, firstData);
    }

    /**
     * remove sheet
     * @param sheet
     */
    public static void removeSheet(Sheet sheet) {
        Workbook workbook = sheet.getWorkbook();
        int sheetIndex = workbook.getSheetIndex(sheet);
        workbook.removeSheetAt(sheetIndex);
    }
    /**
     * Excel area references
     */
    public static class Range {
        /** worksheet */
        private final Sheet sheet;
        /** the distance of the start row of original range and the start row of this range (the cloning range) */
        private final int shiftY;
        /** the distance of the start col of original range and the start col of this range (the cloning range) */
        private final int shiftX;
        /** original range name */
        private final String name;
        private int index;
        private AreaReference areaReference;
        
        public int getIndex() {
			return index;
		}

		public void setIndex(int index) {
			this.index = index;
		}

		public int getShiftY() {
			return shiftY;
		}
        public int getShiftX() {
            return shiftX;
        }

		/**
         * @param sheet the worksheet
         * @param shiftY the shiftY
         * @param shiftX
         * @param name the name
         */
        public Range(Sheet sheet, int shiftY, int shiftX, String name) {
            this.sheet = sheet;
            this.shiftY = shiftY;
            this.shiftX = shiftX;
            this.name = name;
        }

        /**
         * @param sheet the worksheet
         * @param name the name
         */
        public Range(Sheet sheet, String name) {
            this.sheet = sheet;
            this.shiftX = 0;
            this.shiftY = 0;
            this.name = name;
        }

        public AreaReference getAreaReference() {
            if (this.areaReference == null) {
                this.areaReference = new AreaReference(sheet.getWorkbook().getName(name).getRefersToFormula(), null);
            }
            return this.areaReference;
        }

        /**
         * get cell by name in this range
         *
         * @param name
         * @return
         */
        public Cell cell(String name) {
            return  FileUtil.cell(this.sheet, name, this.shiftY, this.shiftX);
        }

        /**
         * copy a range in vertical
         *
         * @param addOffsetY add offset row
         * @return the cloning range
         */
        public Range verticalCopy(int addOffsetY) {
            return this.verticalCopy(this.sheet, addOffsetY, false);
        }

        /**
         * copy a range in vertical
         *
         * @param addOffsetY add offset row
         * @return the cloning range
         */
        public Range verticalCopyInsert(int addOffsetY) {
            return this.verticalCopy(this.sheet, addOffsetY, true);
        }

        /**
         * copy a range in vertical to row index
         *
         * @param sheetDest the destination sheet
         * @param rowIndexTo the destination row index
         * @return the cloning range
         */
        public Range verticalCopyTo(Sheet sheetDest, int rowIndexTo) {
            AreaReference area = this.getAreaReference();
            int lastRow = area.getLastCell().getRow();
            int addOffsetY = rowIndexTo - lastRow - 1;
            return this.verticalCopy(sheetDest, addOffsetY, true);
        }

        /**
         * copy a range in vertical
         * @return the cloning range
         */
        public Range verticalCopy() {
            return verticalCopy(0);
        }
        /**
         * TODO
         * copy a range in horizontal
         * @param addOffsetX add offset row
         * @return the cloning range
         */
        public Range horizontalCopy(int addOffsetX) {
            return this.horizoltalCopy(this.sheet, addOffsetX);
        }

        /**
         * copy a range in horizontal to col index
         *
         * @param sheetDest the destination sheet
         * @param colIndexTo the destination row index
         */
        public Range horizontalCopyTo(Sheet sheetDest, int colIndexTo) {
            AreaReference area = this.getAreaReference();
            int lastCol = area.getLastCell().getCol();
            int addOffsetY = colIndexTo - lastCol - 1;
            return this.horizoltalCopy(sheetDest, addOffsetY);
        }

        /**
         * copy range in vertical
         *
         * @param addOffsetY add offset row
         * @param copyInsert is copy then insert below. in case there're existed rows below, avoid override below content
         * @return the cloning range
         */
        public Range verticalCopy(Sheet sheetDest, int addOffsetY, boolean copyInsert) {
            AreaReference area = this.getAreaReference();
            CellReference firstCell = area.getFirstCell();
            CellReference lastCell = area.getLastCell();
            // row count of original range
            int rowCount = lastCell.getRow() - firstCell.getRow() + 1;
            // the shiftY of start row of the original range and start row of the cloning range
            int shift = this.shiftY + addOffsetY + rowCount;
            // Shifts below rows before copy row down
            if (copyInsert) {
                sheet.shiftRows(lastCell.getRow() + this.shiftY + addOffsetY + 1, sheet.getLastRowNum(), rowCount);
            }
            // copy cell by cell
            for (int y = firstCell.getRow(); y <= lastCell.getRow(); y++) {
                for (int x = firstCell.getCol(); x <= lastCell.getCol(); x++) {
                    Cell srcCell = FileUtil.cell(this.sheet, y, x);

                    // create cloning row when it doesn't existed
                    if (sheetDest.getRow(shift + y) == null) {
                        sheetDest.createRow(shift + y);
                    }
                    Row cloneRow = sheetDest.getRow(shift + y);
                    cloneRow.setHeight(this.sheet.getRow(y).getHeight());

                    Cell cloneCell = cloneRow.createCell(x);
                    FileUtil.copyCell(srcCell, cloneCell);
                }
            }

            // copy merge regions
            for (CellRangeAddress srcRegion : this.sheet.getMergedRegions()) {
                if (firstCell.getRow() <= srcRegion.getFirstRow() && srcRegion.getLastRow() <= lastCell.getRow()) {
                    // srcRegion is fully inside the copied rows
                    final CellRangeAddress destRegion = srcRegion.copy();
                    destRegion.setFirstRow(destRegion.getFirstRow() + shift);
                    destRegion.setLastRow(destRegion.getLastRow() + shift);
                    sheetDest.addMergedRegion(destRegion);
                }
            }
            return new Range(sheetDest, shift, shiftX, this.name);
        }

        /**
         * copy range in horizontal
         *
         * @param addOffsetY add offset row
         * @return the cloning range
         */
        public Range horizoltalCopy(Sheet sheetDest, int addOffsetY) {
            AreaReference area = this.getAreaReference();
            CellReference firstCell = area.getFirstCell();
            CellReference lastCell = area.getLastCell();
            // column count of original range
            int colCount = lastCell.getCol() - firstCell.getCol() + 1;
            // the shiftY of start column of the original range and start column of the cloning range
            int shift = this.shiftX + addOffsetY + colCount;
            // copy cell by cell
            for (int y = firstCell.getRow(); y <= lastCell.getRow(); y++) {
                for (int x = firstCell.getCol(); x <= lastCell.getCol(); x++) {
                    Cell srcCell = FileUtil.cell(this.sheet, y, x);

                    Row cloneRow = sheetDest.getRow(y);

                    Cell cloneCell = cloneRow.createCell(x + shift);
                    FileUtil.copyCell(srcCell, cloneCell);
                }
            }

            // copy merge regions
            for (CellRangeAddress srcRegion : this.sheet.getMergedRegions()) {
                if (firstCell.getCol() <= srcRegion.getFirstColumn() && srcRegion.getLastColumn() <= lastCell.getCol()
                    && firstCell.getRow() <= srcRegion.getFirstRow() && srcRegion.getLastRow() <= lastCell.getRow()) {
                    // srcRegion is fully inside the copied rows
                    final CellRangeAddress destRegion = srcRegion.copy();
                    destRegion.setFirstColumn(destRegion.getFirstColumn() + shift);
                    destRegion.setLastColumn(destRegion.getLastColumn() + shift);
                    sheetDest.addMergedRegion(destRegion);
                }
            }
            return new Range(sheetDest, shiftY, shift, this.name);
        }
    }

    /**
     * replace invalid characters in raw sheet name
     * @param rawName
     * @return
     */
    public static String replaceSheetNameInvalidChar(String rawName) {
        return rawName.replaceAll(EXCEL_NAME_REGEX_INVALID_CHAR, "-");
    }
}

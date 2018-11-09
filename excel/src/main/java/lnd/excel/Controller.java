package lnd.excel;

import lnd.excel.data.Item;
import lnd.excel.data.Supplier;
import lnd.excel.functioninterface.C;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.servlet.annotation.WebServlet;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.net.URISyntaxException;
import java.util.ArrayList;
import java.util.List;

/**
 * @author linhnguyendinh
 */
@WebServlet(value = "/test")
public class Controller extends HttpServlet {
    /**
     *
     */
    private static final long serialVersionUID = 1L;

    public void doGet(HttpServletRequest request, HttpServletResponse response) {
        try {
            List<Item> items = initDataRow(30);
            List<Supplier> sups = initDataCol(20);
            // change to test.xls for xls sample
            this.downloadExcel(response, "test.xlsx", "test.xlsx", w -> {
                Sheet sheet = w.getSheetAt(0);
                // copy down row
                FileUtil.verticalCopyInsertRange(sheet, "row", 0, (range, item) -> {
                    range.cell("itemRef").setCellValue(item.getItemRef());
                    range.cell("desc").setCellValue(item.getDesc());
                    range.cell("quantity").setCellValue(item.getQuatity());
                }, items);
                // copy to the right
                FileUtil.horizontalCopyRange(sheet, "col", 0, (range, sup) -> {
                    range.cell("unitPrice").setCellValue(sup.getUnitPrice());
                    range.cell("totalAmount").setCellValue(sup.getTotalAmount());
                    range.cell("offer").setCellValue(sup.getOffer());
                    range.cell("sampleSubmited").setCellValue(sup.getSampleSubmited());
                    range.cell("remarks").setCellValue(sup.getRemarks());
                }, sups);
            });

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    /**
     * download excel
     *
     * @param response HttpResponse
     * @param templateName the template name
     * @param responseName the response name
     * @param consumer handler: write data to workbook before download
     * @throws Exception
     */
    public void downloadExcel(HttpServletResponse response, String templateName, String responseName, C<Workbook> consumer) throws Exception {
        Workbook workbook = null;
        try {
            File templateFile = new File(this.getClass().getClassLoader().getResource(templateName).toURI());
            FileInputStream inputStream = new FileInputStream(templateFile);
            if (templateName.endsWith("xlsx")) {
                workbook = new XSSFWorkbook(inputStream);
            } else if (templateName.endsWith("xls")) {
                workbook = new HSSFWorkbook(inputStream);
            } else {
                throw new Exception("wrong template file type, file name: " + templateName);
            }

            // handler: write data to workbook
            consumer.accept(workbook);

            response.setContentType("application/pdf");
            response.setHeader("Content-Disposition", "attachment; filename=" + responseName);
            workbook.write(response.getOutputStream()); // Write workbook to response.
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (URISyntaxException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            try {
                if (workbook != null) workbook.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }

    /**
     * init items data (rows data)
     *
     * @param numOfRows
     * @return
     */
    private List<Item> initDataRow(int numOfRows) {
        List<Item> items = new ArrayList<>();
        for (int i = 0; i < numOfRows; i++) {
            Item item = new Item();
            item.setItemRef("[Item Ref.:\n" + i);
            item.setDesc("Electrical Ceiling Fans,complete with fan " + i);
            item.setQuatity(i);
            items.add(item);
        }
        return items;
    }

    /**
     * init suppliers data (cols data)
     *
     * @param numOfCols
     * @return
     */
    private List<Supplier> initDataCol(int numOfCols) {
        List<Supplier> items = new ArrayList<>();
        for (int i = 0; i < numOfCols; i++) {
            Supplier sup = new Supplier();
            sup.setUnitPrice(i);
            sup.setTotalAmount(i * 3);
            sup.setOffer(i % 2 == 0? "N": "Y");
            sup.setSampleSubmited(i % 2 == 0? "Y": "N");
            sup.setRemarks("Remarks " + i);
            items.add(sup);
        }
        return items;
    }
}
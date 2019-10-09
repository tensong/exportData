package exportdata;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.ArrayList;
import java.util.List;
import java.util.Scanner;

/**
 * 携程映射数据脚本工具类
 */
public class ReadCtripDltExcel {

    public static void main(String[] args) {
        Scanner sc = new Scanner(System.in);
        String examplePath = "D:\\ctripTemplate11-16.xlsx";
        System.out.println("please input the mapping file path(eg:" + examplePath + ") end with ENTER");
        String path = sc.nextLine();
        path = path.replaceAll("'|\"", "");
        String fileSeparator = File.separator;
        String defaultExampleStorePath;
        if (path.contains(fileSeparator)) {
            defaultExampleStorePath = path.substring(0, path.lastIndexOf(fileSeparator)) +
                    path.substring(path.lastIndexOf(fileSeparator), path.lastIndexOf(".")) + ".sql";
        } else {
            defaultExampleStorePath = path.substring(0, path.lastIndexOf(".") + 1) + ".sql";
        }

        System.out.println("Start to transform......");

        String lineSeparator = System.getProperty("line.separator");
        // 房仓的价格计划id
        String fcPricePlanId;
        // 携程酒店id
        String ctripHotelId;
        // 携程房型id
        String ctripRoomTypeId;

        try (FileWriter fw = new FileWriter(defaultExampleStorePath, true);
             BufferedWriter bw = new BufferedWriter(fw)) {
            List<List<String>> result = readXls(path);
            for (int i = 0; i < result.size(); i++) {
                List<String> model = result.get(i);
                if (model.size() < 3 || (null == model.get(3) || model.get(3).trim().equals(""))) {
                    continue;
                }

                ctripHotelId = model.get(1).trim();
                ctripRoomTypeId = model.get(3).trim();
                fcPricePlanId = model.get(2).trim();

                String sql3 = "insert into htl_delivery.t_ctrip_dlt_map_rateplan (ID, HOTELID, ROOMTYPEID, ROOMTYPENAME, SUPPLYCODE, PRICEPLANID, PRICEPLANNAME, COMMODITYID, CTRIP_HOTEL_ID, CTRIP_ROOMTYPEID, CTRIP_ROOMTYPENAME, MERCHANTCODE, SUPPLIER_ID, ISACTIVE,createdate,creator)" +
                        "select htl_delivery.SEQ_t_ctrip_dlt_map_rateplan.Nextval,s.hotelid,p.roomtypeid,null,p.supplycode,p.priceplanid,p.priceplanname,s.commodityid,"
                        + "'" + ctripHotelId + "','" + ctripRoomTypeId + "',null,'M10002623',9834,1,sysdate,'sync_181026' from htlpro.T_HTLPRO_COMMODITY s,"
                        + "htlpro.T_HTLPRO_RELATION r, htlpro.T_HTLPRO_PRICEPLAN P where s.COMMODITYID = r.COMMODITYID AND r.ISACTIVE = 1 AND s.PRICEPLANID = P .PRICEPLANID and s.merchantcode = 'M10002623' "
                        + "AND P .priceplanid IN (" + fcPricePlanId + ");";

                if (i != 0 && i % 20 == 0) {
                    bw.append(lineSeparator);
                }

                bw.append(sql3 + lineSeparator);
            }

            Thread.sleep(5000);

            System.out.println("=================SUCCESS!==================");
            System.out.println("=================Result SQL File Path===========");
            System.out.println(defaultExampleStorePath);

        } catch (Exception e) {
            System.out.println("====================MAPPING ERROR!================");
            e.printStackTrace();
        }

    }

    private static List<List<String>> readXls(String path) throws Exception {
        InputStream is = new FileInputStream(path);
        // XSSFWorkbook 标识整个excel
        XSSFWorkbook XSSFWorkbook = new XSSFWorkbook(is);
        List<List<String>> result = new ArrayList<>();
        int size = XSSFWorkbook.getNumberOfSheets();
        // 循环每一页，并处理当前循环页
        for (int numSheet = 0; numSheet < size; numSheet++) {
            // XSSFSheet 标识某一页
            XSSFSheet XSSFSheet = XSSFWorkbook.getSheetAt(numSheet);
            if (XSSFSheet == null) {
                continue;
            }
            // 处理当前页，循环读取每一行
            for (int rowNum = 1; rowNum <= XSSFSheet.getLastRowNum(); rowNum++) {
                // XSSFRow表示行
                XSSFRow XSSFRow = XSSFSheet.getRow(rowNum);
                int minColIx = XSSFRow.getFirstCellNum();
                int maxColIx = XSSFRow.getLastCellNum();
                List<String> rowList = new ArrayList<String>();
                // 遍历改行，获取处理每个cell元素
                for (int colIx = minColIx; colIx < maxColIx; colIx++) {
                    // XSSFCell 表示单元格
                    XSSFCell cell = XSSFRow.getCell(colIx);
                    if (cell == null) {
                        continue;
                    }
                    rowList.add(getStringVal(cell));
                }
                result.add(rowList);
            }
        }
        return result;
    }

    public static String getStringVal(XSSFCell cell) {
        switch (cell.getCellType()) {
            case Cell.CELL_TYPE_BOOLEAN:
                return cell.getBooleanCellValue() ? "TRUE" : "FALSE";
            case Cell.CELL_TYPE_FORMULA:
                return cell.getCellFormula();
            case Cell.CELL_TYPE_NUMERIC:
                cell.setCellType(Cell.CELL_TYPE_STRING);
                return cell.getStringCellValue();
            case Cell.CELL_TYPE_STRING:
                return cell.getStringCellValue();
            default:
                return "";
        }
    }

}

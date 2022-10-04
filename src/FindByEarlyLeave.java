import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.lang.reflect.Field;
import java.time.LocalTime;
import java.time.format.DateTimeFormatter;
import java.util.*;
import java.util.stream.Collectors;

import static java.lang.Thread.sleep;

public class FindByEarlyLeave {

    static DateTimeFormatter timeFormatter = DateTimeFormatter.ofPattern("HH시mm분ss초");
    static DateTimeFormatter dateFormatter = DateTimeFormatter.ofPattern("yyyy년MM월dd일");

    public static void main(String[] args) throws Exception {
        ArrayList<ExcelHeader> excelHeaderDataList = new ArrayList<ExcelHeader>();
        readExcelFile(excelHeaderDataList);
        List<List<ExcelHeader>> groupingDate = processor(excelHeaderDataList);
        writeExcelFile(groupingDate);
    }

    private static void writeExcelFile(List<List<ExcelHeader>> groupingCompany) throws IllegalAccessException {
        for(List<ExcelHeader> dataList : groupingCompany){
            String excelName = "None";
            XSSFWorkbook workbook = new XSSFWorkbook();
            XSSFSheet sheet = workbook.createSheet("sheet");
            sheet.setDefaultColumnWidth(15);

            int rownum = 0;
            for (ExcelHeader data : dataList) {
                Row row = sheet.createRow(rownum++);
                excelName = data.getCompanyName();
                int cellIndex = 0;
                Cell cell0 = row.createCell(0);
                Cell cell1 = row.createCell(1);
                Cell cell2 = row.createCell(2);
                Cell cell3= row.createCell(3);
                Cell cell4= row.createCell(4);
                Cell cell5= row.createCell(5);
                cell0.setCellValue(data.getFormatIssueDate());
                cell1.setCellValue(data.getFormatIssueTime());
                cell2.setCellValue(data.getCardNumber());
                cell3.setCellValue(data.getUserName());
                cell4.setCellValue(data.getCompanyName());
                cell5.setCellValue(data.getDeviceName());
            }
            try (FileOutputStream out = new FileOutputStream(new File("./", excelName + ".xlsx"))) {
                workbook.write(out);
                workbook.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }

    private static List<List<ExcelHeader>> processor(ArrayList<ExcelHeader> excelHeaderDataList) {
        Comparator<ExcelHeader> compare = Comparator
                .comparing(ExcelHeader::getCardNumber)
                .thenComparing(ExcelHeader::getIssueDate);

        List<List<ExcelHeader>> groupingData = excelHeaderDataList.stream()
                .filter(x -> isEarlyLeaveTime(x.getIssueTime()))
                .collect(Collectors.groupingBy(ExcelHeader::getCompanyName))
                .entrySet()
                .stream()
                .map(x -> x.getValue().stream().sorted(compare).collect(Collectors.toList()))
                .collect(Collectors.toList());
        return groupingData;
    }

    private static void readExcelFile(ArrayList<ExcelHeader> excelHeaderDataList) throws Exception {
        try {
            //파일 read
            File readFilePath = new File("./INPUT.xlsx");
            Workbook workbook = new XSSFWorkbook(new FileInputStream(readFilePath));

            Sheet worksheet = workbook.getSheetAt(0);
            Iterator<Row> rowIterator = worksheet.iterator();

            //첫번쨰 행 Skip
            rowIterator.next();
            while (rowIterator.hasNext()) {
                Row readRow = rowIterator.next();
                Iterator<Cell> cellIterator = readRow.cellIterator();
                ExcelHeader excelHeaderData = new ExcelHeader();

                List<String> values = new ArrayList<>();
                List<String> colums = new ArrayList<>();
                String cellData;
                while (cellIterator.hasNext()) {
                    Cell readCell = cellIterator.next();

                    //컬럼 type 분기처리
                    switch (readCell.getCellType()) {
                        case STRING:
                            cellData = readCell.getRichStringCellValue().getString();
                            break;
                        case NUMERIC:
                            if (DateUtil.isCellDateFormatted(readCell)) {
                                cellData = readCell.getDateCellValue().toString();
                            } else {
                                Long roundVal = Math.round(readCell.getNumericCellValue());
                                Double doubleVal = readCell.getNumericCellValue();
                                if (doubleVal.equals(roundVal.doubleValue())) {
                                    cellData = String.valueOf(roundVal);
                                } else {
                                    cellData = String.valueOf(doubleVal);
                                }
                            }
                            break;
                        case BOOLEAN:
                            cellData = String.valueOf(readCell.getBooleanCellValue());
                            break;
                        case FORMULA:
                            cellData = readCell.getCellFormula();
                            break;
                        default:
                            cellData = "";
                    }

                    int cellIndex = readCell.getColumnIndex();
                    excelHeaderData.setHeaderIndexValue(cellIndex, cellData);
                }

                excelHeaderDataList.add(excelHeaderData);
            }
            workbook.close();

        }  catch (Exception e) {
            throw new Exception(e.getMessage());
        }
    }

    static boolean isEarlyLeaveTime(LocalTime issueTime){
        List<Integer> earlyTime = Arrays.asList(6, 14, 16, 22);
        return earlyTime.indexOf(issueTime.getHour()) >= 0;
    }
}

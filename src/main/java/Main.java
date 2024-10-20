import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

public class Main {
    private static final String FILE_NAME = "C:\\Users\\hahuy\\Desktop\\SWT301_LAB-1_Templates (1).xlsx";

    // Đọc file Excel và trả về danh sách Input
    public List<Input> readExcel() throws IOException {
        List<Input> inputList = new ArrayList<>();
        FileInputStream excelFile = new FileInputStream(FILE_NAME);
        Workbook workbook = new XSSFWorkbook(excelFile);
        Sheet sheet = workbook.getSheetAt(0); // Lấy sheet đầu tiên

        for (int colIndex = 5; colIndex <= 19; colIndex++) { // Cột F (index = 5) đến T (index = 19)
            String a = "";
            String b = "";
            String c = "";

        for (int rowIndex = 14; rowIndex <= 19; rowIndex++) { // Hàng 15 đến 20
            Row row = sheet.getRow(rowIndex);
            if (row != null) {
                Cell cellF = row.getCell(colIndex);
                if (cellF != null && cellF.getCellType() != CellType.BLANK) {
                    Cell cellD = row.getCell(3); // Cột D (index = 3)
                    a = (cellD != null && cellD.getCellType() != CellType.BLANK) ? cellD.toString() : "";
                    break; // Có giá trị, không cần kiểm tra tiếp trong phạm vi này
                }
            }
        }

        for (int rowIndex = 21; rowIndex <= 26; rowIndex++) { // Hàng 22 đến 27
            Row row = sheet.getRow(rowIndex);
            if (row != null) {
                Cell cellF = row.getCell(colIndex);
                if (cellF != null && cellF.getCellType() != CellType.BLANK) {
                    Cell cellD = row.getCell(3); // Cột D (index = 3)
                    b = (cellD != null && cellD.getCellType() != CellType.BLANK) ? cellD.toString() : "";
                    break; // Có giá trị, không cần kiểm tra tiếp trong phạm vi này
                }
            }
        }

        for (int rowIndex = 28; rowIndex <= 31; rowIndex++) { // Hàng 29 đến 32
            Row row = sheet.getRow(rowIndex);
            if (row != null) {
                Cell cellF = row.getCell(colIndex);
                if (cellF != null && cellF.getCellType() != CellType.BLANK) {
                    Cell cellD = row.getCell(3); // Cột D (index = 3)
                    c = (cellD != null && cellD.getCellType() != CellType.BLANK) ? cellD.toString() : "";
                    break; // Có giá trị, không cần kiểm tra tiếp trong phạm vi này
                }
            }
        }

        // Thêm giá trị vào danh sách input
        inputList.add(new Input(a, b, c));
    }



        workbook.close();
        return inputList;
    }


    public void writeExcel(List<Result> resultList) throws IOException {
        FileInputStream excelFile = new FileInputStream(FILE_NAME);
        Workbook workbook = new XSSFWorkbook(excelFile);
        Sheet sheet = workbook.getSheetAt(0);

        // Đảm bảo rằng hàng 51, 42, 43, 39, 40 và 41 tồn tại
        Row row50 = sheet.getRow(49);
        if (row50 == null) {
            row50 = sheet.createRow(49);
        }

        Row row51 = sheet.getRow(50); // Hàng 51 (index = 50)
        if (row51 == null) {
            row51 = sheet.createRow(50);
        }

        Row row42 = sheet.getRow(41); // Hàng 42 (index = 41)
        if (row42 == null) {
            row42 = sheet.createRow(41);
        }

        Row row43 = sheet.getRow(42); // Hàng 43 (index = 42)
        if (row43 == null) {
            row43 = sheet.createRow(42);
        }

        Row row39 = sheet.getRow(38); // Hàng 39 (index = 38)
        if (row39 == null) {
            row39 = sheet.createRow(38);
        }

        Row row40 = sheet.getRow(39); // Hàng 40 (index = 39)
        if (row40 == null) {
            row40 = sheet.createRow(39);
        }

        Row row41 = sheet.getRow(40); // Hàng 41 (index = 40)
        if (row41 == null) {
            row41 = sheet.createRow(40);
        }

        int colIndex = 5; // Bắt đầu từ cột F
        for (Result result : resultList) {
            if (colIndex > 19) {
                break;
            }

            Cell cellType = row50.createCell(colIndex);
            cellType.setCellValue(result.getType());

            // Ghi result.message vào hàng 51
            Cell cellMessage = row51.createCell(colIndex);
            cellMessage.setCellValue(result.getMessage());

            // Ghi result.x1 vào hàng 42
            Cell cellX1 = row42.createCell(colIndex);
            cellX1.setCellValue(result.getX1());

            // Ghi result.x2 vào hàng 43
            Cell cellX2 = row43.createCell(colIndex);
            cellX2.setCellValue(result.getX2());

            // Xác định vị trí cột cho hàng 39, 40 và 41
            Cell cell39 = row39.createCell(colIndex);
            Cell cell40 = row40.createCell(colIndex);
            Cell cell41 = row41.createCell(colIndex);

            // Kiểm tra các điều kiện và ghi giá trị vào hàng 39, 40, 41
            if(result.getType().equals("N")) {
                if (result.getX1() != null && result.getX2() != null && !result.getX1().equals(result.getX2())) {
                    cell39.setCellValue("O");
                } else if (result.getX1() == null && result.getX2() == null) {
                    cell40.setCellValue("O");
                } else {
                    cell41.setCellValue("O");
                }
            }

            colIndex++;
        }

        excelFile.close();

        // Ghi file lại
        FileOutputStream outFile = new FileOutputStream(FILE_NAME);
        workbook.write(outFile);
        outFile.close();
        workbook.close();
    }

    public static void main(String[] args) throws IOException {
        Main handler = new Main();
        Caculate calculator = new Caculate();

        // Đọc dữ liệu từ Excel
        List<Input> inputList = handler.readExcel();
        List<Result> resultList = new ArrayList<>();

        // Tính toán kết quả
        for (Input input : inputList) {
            resultList.add(calculator.Caculate(input));
        }

        // Ghi kết quả vào Excel
        handler.writeExcel(resultList);
    }
}


import java.text.DecimalFormat;

public class Caculate {
    private static final DecimalFormat DECIMAL_FORMAT = new DecimalFormat("#.00");

    public Result Caculate(Input i) {
        Result rs = new Result();
        try {
            // Chuyển đổi các chuỗi số thực thành số nguyên
            int a = (int) Math.round(Double.parseDouble(i.getA()));
            int b = (int) Math.round(Double.parseDouble(i.getB()));
            int c = (int) Math.round(Double.parseDouble(i.getC()));

            // Kiểm tra giá trị hợp lệ
            if (a <= 0 || b <= 0 || c <= 0 || a >= 65535 || b >= 65535 || c >= 65535) {
                rs.message = "F";
                rs.type = "B";
            } else {
                double delta = b * b - 4 * a * c;

                if (delta > 0) {
                    double x1 = ((-b + Math.sqrt(delta)) / (2 * a));
                    double x2 = ((-b - Math.sqrt(delta)) / (2 * a));
                    rs.x1 = DECIMAL_FORMAT.format(x1);
                    rs.x2 = DECIMAL_FORMAT.format(x2);
                    rs.message = "P";
                    rs.type = "N";
                } else if (delta == 0) {
                    double x1 = -b / (2 * a);
                    rs.x1 = DECIMAL_FORMAT.format(x1);
                    rs.x2 = rs.x1;
                    rs.message = "P";
                    rs.type = "N";
                } else {
                    rs.x1 = null;
                    rs.x2 = null;
                    rs.message = "P";
                    rs.type = "N";
                }
            }
        } catch (Exception e) {
            rs.message = "F";
            rs.type = "A";
        }
        return rs;
    }
}

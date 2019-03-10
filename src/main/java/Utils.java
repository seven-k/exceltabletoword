import java.text.SimpleDateFormat;

/**
 * @author yin.cl
 * @since 2019/3/9 11:46
 */
public class Utils {
    private static final String TIME_FORMAT = "yyyyMMddhhmmss";

    public static String getCurrentTimeStr2() {
        long now = System.currentTimeMillis();
        SimpleDateFormat dateFormat = new SimpleDateFormat(TIME_FORMAT);
        return dateFormat.format(now);
    }

    public static boolean isEmpty(Object o) {
        return o == null || "".equals(o);
    }

    public static boolean isNotEmpty(Object o) {
        return !isEmpty(o);
    }
}

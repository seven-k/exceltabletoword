import lombok.Builder;
import lombok.Data;

/**
 * @author yin.cl
 * @since 2019/3/9 11:36
 */
@Data
@Builder
public class User {
    private String userName;
    private String sex;
    private String birthday;
    private String cultureLevel;
    private String phone;
    private String address;

    public static void main(String[] args) {
        String userName="ABC";
        String filePath="/Users/yiyezhiqiu/Desktop/df/个人备案.docx";
        int dotAt = filePath.lastIndexOf(".");
        StringBuilder sb = new StringBuilder(filePath);
        sb.replace(dotAt, dotAt, "_" + userName + "_" + Utils.getCurrentTimeStr2());
        System.out.println(sb.toString());
    }
}

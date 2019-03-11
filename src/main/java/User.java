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

}

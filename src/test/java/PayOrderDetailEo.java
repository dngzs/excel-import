import com.best.excel.annotation.ExcelEo;
import com.best.excel.annotation.ExcelTitle;
import org.hibernate.validator.constraints.NotBlank;

import javax.validation.constraints.Pattern;
import java.math.BigDecimal;
import java.util.Date;

/**
 * @author chenandong
 * @date 2018/7/12 11:35
 */
@ExcelEo("PayOrderDetailEo")
public class PayOrderDetailEo{
    private static final long serialVersionUID = -203167215073199823L;

    @ExcelTitle(name = "罚/返款对象编码")
    @NotBlank(message = "罚/返款对象编码不能为空")
    private String chargeCode;

    @ExcelTitle(name = "罚/返款对象名称")
    @NotBlank(message = "罚/返款对象名称不能为空")
    private String chargeName;

    @ExcelTitle(name = "金额")
    private BigDecimal paidMoney;

    @ExcelTitle(name = "运单号")
    @Pattern(regexp = "^\\s*?[a-zA-Z0-9]+\\s*?$", message = "运单号只能是数字和字母组成")
    private String transOrderCode;

    @ExcelTitle(name = "费用说明",format = "yyyy-MM-dd")
    private Date instruction;

    @Override
    public String toString() {
        return "PayOrderDetailEo{" +
                "chargeCode='" + chargeCode + '\'' +
                ", chargeName='" + chargeName + '\'' +
                ", paidMoney=" + paidMoney +
                ", transOrderCode='" + transOrderCode + '\'' +
                ", instruction='" + instruction + '\'' +
                '}';
    }

    public static long getSerialVersionUID() {
        return serialVersionUID;
    }

    public String getChargeCode() {
        return chargeCode;
    }

    public void setChargeCode(String chargeCode) {
        this.chargeCode = chargeCode;
    }

    public String getChargeName() {
        return chargeName;
    }

    public void setChargeName(String chargeName) {
        this.chargeName = chargeName;
    }

    public BigDecimal getPaidMoney() {
        return paidMoney;
    }

    public void setPaidMoney(BigDecimal paidMoney) {
        this.paidMoney = paidMoney;
    }

    public String getTransOrderCode() {
        return transOrderCode;
    }

    public void setTransOrderCode(String transOrderCode) {
        this.transOrderCode = transOrderCode;
    }

    public Date getInstruction() {
        return instruction;
    }

    public void setInstruction(Date instruction) {
        this.instruction = instruction;
    }
}

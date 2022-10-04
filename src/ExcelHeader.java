import java.time.LocalDate;
import java.time.LocalTime;

public class ExcelHeader {

    private LocalDate issueDate;
    private LocalTime issueTime;
    private String cardNumber;
    private String userName;
    private String companyName;
    private String deviceName;

    public LocalDate getIssueDate() {
        return issueDate;
    }

    public String getFormatIssueDate() {
        return FindByEarlyLeave.dateFormatter.format(issueDate);
    }

    public void setIssueDate(LocalDate issueDate) {
        this.issueDate = issueDate;
    }

    public LocalTime getIssueTime() {
        return issueTime;
    }

    public String getFormatIssueTime() {
        return FindByEarlyLeave.timeFormatter.format(issueTime);
    }

    public void setIssueTime(LocalTime issueTime) {
        this.issueTime = issueTime;
    }

    public String getCardNumber() {
        return cardNumber;
    }

    public void setCardNumber(String cardNumber) {
        this.cardNumber = cardNumber;
    }

    public String getUserName() {
        return userName;
    }

    public void setUserName(String userName) {
        this.userName = userName;
    }

    public String getCompanyName() {
        return companyName;
    }

    public void setCompanyName(String companyName) {
        this.companyName = companyName;
    }

    public String getDeviceName() {
        return deviceName;
    }

    public void setDeviceName(String deviceName) {
        this.deviceName = deviceName;
    }

    public void setHeaderIndexValue(int index, String value){
        if(index == 0){
            this.issueDate = LocalDate.parse(value, FindByEarlyLeave.dateFormatter);
        } else if(index == 1){
            this.issueTime = LocalTime.parse(value, FindByEarlyLeave.timeFormatter);
        } else if(index == 2){
            this.cardNumber = value;
        } else if(index == 3){
            this.userName = value;
        } else if(index == 4){
            this.companyName = value;
        } else if(index == 5){
            this.deviceName = value;
        }
    }
}

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.mail.SimpleMailMessage;
import org.springframework.mail.javamail.JavaMailSenderImpl;

import java.io.*;

/**
 * User: Lichao.W
 * Date: 2014/10/24
 * Time: 16:28
 */
public class SendMailGroup {
    private static final int DEFAULT_PAUSE_TIME_SECONDS = 140;
    private static final String femail = "factoryrubber@163.com";
    private static final String passwd = "*******";
    private String title = "factoryrubber";//默认主题

    public void readXLS() {
        try {
            XSSFWorkbook xwb = new XSSFWorkbook("d:/sendmail/email.xlsx");
            //String ename = ""; //收件人姓名
            String email = ""; //收件人地址
            String tmp = "";
            XSSFSheet sheet = xwb.getSheetAt(0);
            int rowLen = sheet.getLastRowNum();
            XSSFCell cell;
            for (int i = 1; i <= rowLen; i++) {
                cell = sheet.getRow(i).getCell(0);
                email = cell.getStringCellValue().trim();
                sendEmail(email);
                Thread.sleep(1000 * DEFAULT_PAUSE_TIME_SECONDS);
            }
        } catch (InterruptedException e) {
            System.out.println("Thread Error");
            e.printStackTrace();
        }  catch (IOException e) {
            System.out.println("read excel error");
            e.printStackTrace();

        }
    }

    public String readTxt() {
        File file = new File("d:/sendmail/email.txt");
        BufferedReader reader = null;
        String tempString = null;
        StringBuilder sb = new StringBuilder();
        int line = 1;
        try {
            reader = new BufferedReader(new FileReader(file));
            while ((tempString = reader.readLine()) != null) {
                if (line == 1) {
                    title = tempString;
                } else {
                    sb.append(tempString + "\n");
                }
                line++;
            }
        } catch (FileNotFoundException e) {
            System.out.println("文件不存在....");
            e.printStackTrace();
        } catch (IOException e) {
            System.out.println("文件读取错误....");
            e.printStackTrace();
        } finally {
            try {
                if (reader != null) {
                    reader.close();
                }
            } catch (IOException e) {
                throw new RuntimeException(e);
            }
        }
        return sb.toString();
    }

    public void sendEmail(String email) {
        JavaMailSenderImpl senderImpl = new JavaMailSenderImpl();
        senderImpl.setHost("smtp.163.com");
        senderImpl.setUsername(femail);
        senderImpl.setPassword(passwd);
        String messg = readTxt();
        SimpleMailMessage mailMessage = new SimpleMailMessage();
        mailMessage.setTo(email);
        mailMessage.setFrom(femail);
        mailMessage.setSubject(title);
        mailMessage.setText(messg);
        senderImpl.send(mailMessage);

        System.out.println("邮件发送成功...to..." + email);
    }

    public static void main(String[] args) {
        SendMailGroup send = new SendMailGroup();
        send.readXLS();
    }
}

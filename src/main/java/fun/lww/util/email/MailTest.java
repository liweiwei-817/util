package fun.lww.util.email;

public class MailTest {

    public static void main(String[] args) throws Exception {
        MailSender mailSender = MailSender.getInstance();
        MailInfo mailInfo = mailSender.getMailInfo();
        mailInfo.setNotifyTo("2622026762@qq.com");//收件人
        mailInfo.setNotifyCc("2622026762@qq.com");//抄送人
        mailInfo.setSubject("123");//主题
        mailInfo.setContent("123123123");//内容
        mailSender.sendHtmlMail(mailInfo, 3);
    }
}

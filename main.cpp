#include <QAxObject>
#include <QCoreApplication>
#include <QSqlDatabase>
#include <QSqlTableModel>
#include <QSqlError>
#include <QSqlQuery>
#include <QFile>
#include <QSettings>
#include <QLocale>
#include <QThread>
#include <src/SmtpMime>


#include <QDebug>

int main(int argc, char *argv[])
{
    QCoreApplication a(argc, argv);


    QFile st(QCoreApplication::applicationDirPath()+"/settings.ini");
    qDebug()<<"DIR1:"<<QCoreApplication::applicationDirPath();
    QString settingsFile=QCoreApplication::applicationDirPath() +"/settings.ini";
    QSettings settings(settingsFile,QSettings::IniFormat);

    if (!st.open(QIODevice::ReadOnly | QIODevice::Text))
    {

        settings.setValue("dsnMySQL","serviceteam");
        settings.setValue("hostnameMySQL","localhost");
        settings.setValue("templateFullPath","c:\\jim\\template.xls");
        settings.setValue("excelFullPath","c:\\jim\\schedule.xls");
        settings.setValue("smtpHostname","smtp.gmail.com");
        settings.setValue("smtpPort",465);
        settings.setValue("smtpUser","algosakis@gmail.com");
        settings.setValue("smtpPass","v@$230698");
        settings.setValue("sender","algosakis@gmail.com");
        settings.setValue("senderName","Sakis");
        settings.setValue("recipient","jimpar@algo.gr");
        settings.setValue("recipientName","jimpar@algo.gr");
        settings.setValue("subject","Schedule for ");
        settings.setValue("body","Schedule for ");
        settings.sync();
    }
    st.close();

    QFile file1(settings.value("templateFullpath").toString());

    file1.copy(settings.value("excelFullpath").toString());

    QSqlDatabase db1=QSqlDatabase::addDatabase("QODBC3",settings.value("dsnMySQL").toString());
    db1.setDatabaseName(settings.value("dsnMySQL").toString());
    db1.setHostName(settings.value("hostnameMySQL").toString());
    if (db1.open()==false)
    {
        qDebug()<<"Error on MySQL:"<< db1.lastError().text();
        exit(0);
    }

    QSqlDatabase db3=QSqlDatabase::addDatabase("QODBC3");


    db3.setDatabaseName("DRIVER={Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)};DBQ="+settings.value("excelFullPath").toString()+";MODE=ReadWrite;Readonly=false");
    if (db3.open()==false)
    {
        qDebug()<<"Error on Excel:"<< db3.lastError().text();
        exit(0);
    }
    QSqlQuery query1(db1),query2(db1);
    QSqlQuery query3(db3);
    setlocale(LC_ALL,"");

    query1.exec("select distinct(ifnull(left(c.county,3),'-')) from customer c,ticket t\
                where t.cusid=c.id and t.appointmentfrom>now() and  t.appointmentfrom<now()+\
                interval 1 day");


    while (query1.next())
    {
            int i=0;

            query3.exec("create table "+query1.value(0).toString()+" (customer varchar,\
                        address varchar,location varchar,city varchar,county varchar,\
                        phone1 varchar,phone2 varchar,loopnumber varchar,company varchar,\
                        postalcode varchar,title varchar,reporteddate varchar,priority varchar,\
                        description varchar,customerticketno varchar,appointmentfrom varchar,\
                        incident varchar,department varchar)");


                    QString qry="select c.name,c.address,c.location,c.city,c.county,\
                    c.phone1,c.phone2,c.loopnumber,o.name,c.postalcode,t.title,t.reporteddate,t.priority,\
                    t.customerticketno,t.appointmentfrom,t.incident,t.department,t.description from customer c,\
                    ticket t,originator o where t.cusid=c.id and o.id=t.originatorid and left(c.county,3)='"+query1.value(0).toString()+"' and t.appointmentfrom>now()\
                    and  t.appointmentfrom<now()+ interval 1 day";


            query2.exec(qry);

              while (query2.next())


{

                   qDebug()<<query2.value(0).toString();

                    query3.exec("insert into "+query1.value(0).toString()+" (customer,address,location,city,county,\
                                phone1,phone2,loopnumber,company,postalcode,title,reporteddate,priority,\
                                customerticketno,appointmentfrom,incident,department,description) VALUES ('"+query2.value(0).toString()+"','"+\
                                query2.value(1).toString()+"','"+query2.value(2).toString()+"','"+query2.value(3).toString()+"','"+\
                                query2.value(4).toString()+"','"+query2.value(5).toString()+"','"+query2.value(6).toString()+"','"+\
                                query2.value(7).toString()+"','"+query2.value(8).toString()+"','"+query2.value(9).toString()+"','"+\
                                query2.value(10).toString()+"','"+query2.value(11).toString()+"','"+query2.value(12).toString()+"','"+\
                                query2.value(13).toString()+"','"+query2.value(14).toString()+"','"+query2.value(15).toString()+"','"+\
                                query2.value(16).toString()+"','"+query2.value(17).toString()+"')");

    }







}

//-----------------------ALL RECORDS


    query3.exec("create table ΟΛΑ (customer varchar,\
                address varchar,location varchar,city varchar,county varchar,\
                phone1 varchar,phone2 varchar,loopnumber varchar,company varchar,\
                postalcode varchar,title varchar,reporteddate varchar,priority varchar,\
                description varchar,customerticketno varchar,appointmentfrom varchar,\
                incident varchar,department varchar)");


            QString qry="select c.name,c.address,c.location,c.city,c.county,\
            c.phone1,c.phone2,c.loopnumber,o.name,c.postalcode,t.title,t.reporteddate,t.priority,\
            t.customerticketno,t.appointmentfrom,t.incident,t.department,t.description from customer c,\
            ticket t,originator o where t.cusid=c.id and o.id=t.originatorid  and t.appointmentfrom>now()\
            and  t.appointmentfrom<now()+ interval 1 day";

            query2.exec(qry);
            qDebug()<<qry;





      while (query2.next())


{


          query3.exec("insert into ΟΛΑ (customer,address,location,city,county,\
                      phone1,phone2,loopnumber,company,postalcode,title,reporteddate,priority,\
                      customerticketno,appointmentfrom,incident,department,description) VALUES ('"+query2.value(0).toString()+"','"+\
                      query2.value(1).toString()+"','"+query2.value(2).toString()+"','"+query2.value(3).toString()+"','"+\
                      query2.value(4).toString()+"','"+query2.value(5).toString()+"','"+query2.value(6).toString()+"','"+\
                      query2.value(7).toString()+"','"+query2.value(8).toString()+"','"+query2.value(9).toString()+"','"+\
                      query2.value(10).toString()+"','"+query2.value(11).toString()+"','"+query2.value(12).toString()+"','"+\
                      query2.value(13).toString()+"','"+query2.value(14).toString()+"','"+query2.value(15).toString()+"','"+\
                      query2.value(16).toString()+"','"+query2.value(17).toString()+"')");

}
    db3.close();
    db1.close();


            SmtpClient smtp(settings.value("smtpHostname").toString(), settings.value("smtpPort").toInt(), SmtpClient::SslConnection);

            // We need to set the username (your email address) and the password
            // for smtp authentification.

            smtp.setUser(settings.value("smtpUser").toString());
            smtp.setPassword(settings.value("smtpPass").toString());

            // Now we create a MimeMessage object. This will be the email.

            MimeMessage message;

            message.setSender(new EmailAddress(settings.value("sender").toString(),settings.value("senderName").toString()));
            message.addRecipient(new EmailAddress(settings.value("recipient").toString(),settings.value("recipientName").toString()));
            message.setSubject(settings.value("subject").toString()+QDate::currentDate().addDays(1).toString("dd.MM.yyyy"));
            message.addPart(new MimeAttachment(new QFile(settings.value("excelFullPath").toString())));

            // Now add some text to the email.
            // First we create a MimeText object.

            MimeText text;

            text.setText(settings.value("body").toString()+QDate::currentDate().addDays(1).toString("dd.MM.yyyy"));

            // Now add it to the mail

            message.addPart(&text);

            // Now we can send the mail

            smtp.connectToHost();
            smtp.login();
            smtp.sendMail(message);
            smtp.quit();
            QFile file2(settings.value("excelFullpath").toString());
            file2.remove();





      QThread::msleep(35000);
      exit (0);

//    return a.exec();
}

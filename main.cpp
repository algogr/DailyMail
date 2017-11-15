#include <QAxObject>
#include <QCoreApplication>
#include <QSqlDatabase>
#include <QSqlTableModel>
#include <QSqlError>
#include <QSqlQuery>
#include <QFile>
#include <QSettings>
#include <QLocale>

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
        settings.setValue("dsnSQLServer","soft1");
        settings.setValue("userSQLServer","sa");
        settings.setValue("passSQLServer","tt123!@#");
        settings.setValue("excelFullPath","c:\\jim\\schedule.xls");
        settings.setValue("batchFilesPath","C:\\algo\\AutoImport\\");
        settings.setValue("cusBatchFilename","AutoRunCusImport.bat");
        settings.setValue("salBatchFilename","AutoRunSalImport.bat");
        settings.setValue("imapHostname","imap.gmail.com");
        settings.setValue("imapPort",993);
        settings.setValue("imapUser","algosakis@gmail.com");
        settings.setValue("imapPass","v@$230698");
        settings.sync();
    }
    st.close();

    QSqlDatabase db1=QSqlDatabase::addDatabase("QODBC3",settings.value("dsnMySQL").toString());
    db1.setDatabaseName(settings.value("dsnMySQL").toString());
    db1.setHostName(settings.value("hostnameMySQL").toString());
    if (db1.open()==false)
    {
        qDebug()<<"Error on MySQL:"<< db1.lastError().text();
        exit(0);
    }

    QSqlDatabase db3=QSqlDatabase::addDatabase("QODBC3");
    //db3.setDatabaseName("DRIVER={Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)};DBQ="+QString(settings.value("excelFullPath").toString()));
    //db3.setDatabaseName("DRIVER={Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)};DBQ=c:\\jim\\jim1.xlsx;MODE=ReadWrite;Readonly=false;Extended Properties-\"HDR\"='YES'");
    db3.setDatabaseName("DRIVER={Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)};DBQ=c:\\jim\\1.xls;MODE=ReadWrite;Readonly=false");
    if (db3.open()==false)
    {
        qDebug()<<"Error on Excel:"<< db3.lastError().text();
        exit(0);
    }
/*
    else
    {
        qDebug()<<db3.isOpen();
        QSqlQuery query(db3);


        query.exec("select * from [Φύλλο1$]");
        while (query.next())
        {
            qDebug()<<query.value(0).toString();
        }

        query.exec("create table mytable (m3 varchar,m4 varchar)");

        query.exec("create table mytable2 (m1 varchar,m2 varchar)");

        qDebug()<<"Excel:"<< db3.lastError().text();

        query.exec("insert into mytable2 (m1 ,m2) values ('test','test1')");

        qDebug()<<"Excel:"<< db3.lastError().text();
        query.exec("DROP TABLE mytable1");
        db3.close();

        exit(0);
    }

*/

    QSqlQuery query1(db1),query2(db1);
    QSqlQuery query3(db3);
    setlocale(LC_ALL,"");

    query1.exec("select distinct(ifnull(left(c.county,3),'-')) from customer c,ticket t\
                where t.cusid=c.id and t.appointmentfrom>now() and  t.appointmentfrom<now()+\
                interval 1 day");

            qDebug()<<"SIZE:"<<query1.size();
    while (query1.next())
    {
            qDebug()<<query1.value(0).toString();

            query3.exec("create table "+query1.value(0).toString()+" (customer varchar,\
                        address varchar,location varchar,city varchar,county varchar,\
                        phone1 varchar,phone2 varchar,loopnumber varchar,company varchar,\
                        postalcode varchar,title varchar,reporteddate varchar,priority varchar,\
                        description varchar,customerticketno varchar,appointmentfrom varchar,\
                        incident varchar,department varchar)");


                    QString qry="select c.name,c.address,c.location,c.city,c.county,\
                    c.phone1,c.phone2,c.loopnumber,o.name,c.postalcode,t.title,t.reporteddate,t.priority,\
                    t.customerticketno,t.appointmentfrom,t.incident,t.department from customer c,\
                    ticket t,originator o where t.cusid=c.id and o.id=t.originatorid and left(c.county,3)='"+query1.value(0).toString()+"' and t.appointmentfrom>now()\
                    and  t.appointmentfrom<now()+ interval 1 day";
                    qDebug()<<qry;
            query2.exec(qry);

              while (query2.next())

                    qDebug()<<query2.value(0).toString();
{

                    query3.exec("insert into "+query1.value(0).toString()+" (customer,address,location,city,county,\
                                phone1,phone2,loopnumber,company,postalcode,title,reporteddate,priority,\
                                customerticketno,appointmentfrom,incident,department) VALUES ('"+query2.value(0).toString()+"','"+\
                                query2.value(1).toString()+"','"+query2.value(2).toString()+"','"+query2.value(3).toString()+"','"+\
                                query2.value(4).toString()+"','"+query2.value(5).toString()+"','"+query2.value(6).toString()+"','"+\
                                query2.value(7).toString()+"','"+query2.value(8).toString()+"','"+query2.value(9).toString()+"','"+\
                                query2.value(10).toString()+"','"+query2.value(11).toString()+"','"+query2.value(12).toString()+"','"+\
                                query2.value(13).toString()+"','"+query2.value(14).toString()+"','"+query2.value(15).toString()+"','"+\
                                query2.value(16).toString()+"')");

    }







}


            db3.close();
            db1.close();

exit(0);



    return a.exec();
}

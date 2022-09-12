drop table tMSP
go 
create table tMSP
(INN_IP varchar(20)
 ,INN_UL varchar(20)
 ,DateMSP varchar(10)
 ,KatMSP int 
 ,FileID varchar(1000)
 )
 go 
 grant all on tMSP to public

--select * from tMSP where INN_IP = '7707794576'
--delete tMSP
--select count(1) from tMSP
--type = "xs:date"
drop table pChangeMSP
go
create table pChangeMSP (FileID varchar(255), InstitutionID DSIDENTIFIER, INN varchar(20), KatMSP int, Datestart smalldatetime, DateEnd smalldatetime, fl int,RetVal int)
go
grant all on pChangeMSP to public
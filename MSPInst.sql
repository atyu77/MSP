create proc MSPInst @date smalldatetime
as

declare @date1  smalldatetime,        
        @date2  smalldatetime,
		@InstID DSIDENTIFIER,
		@RMSPInstitutionID DSIDENTIFIER,
		@ID DSIDENTIFIER,
		@KatMSP int,
		@Dt     smalldatetime,
		@Comment varchar(255),
		@RetVal  int


select @date1 =EOMONTH(dateadd(mm,-2,@date)) 
       ,@date2 = EOMONTH(dateadd(mm,-1,@date))
       ,@RetVal = 0
                
delete pResourceParm where spid=@@spid
delete pRestCalendarOut where spid=@@spid

insert pResourceParm(spid,ResourceID,AccNumber)
select @@spid
       ,ResourceID
       ,''
  from tResource    r WITH(NOLOCK)
  join tInstitution i WITH(NOLOCK) on i.InstitutionID = r.InstOwnerID 
  join tInstAttr   ia WITH(NOLOCK) on i.InstitutionID = ia.InstitutionID 
 where 1=1
   and r.ResourceType = 1
   and r.BalanceID = 2140
   and r.InstitutionID = 2000
   and left(r.Brief,3) in ('407','408')
   and left(r.Brief,5) not in ('40807')
   and (isnull(DateEnd,'19000101') = '19000101' or DateEnd >= @date1) 
   and ((i.PropDealPart = 1)
    or (i.PropDealPart = 0) and (convert(int, ia.RegTemp)&32)/32=1 
           )
   

exec RestByCalendar2 @CalcRestAlg = 2,
                      @BegDate = @Date1,
                      @EndDate = @Date2,
                      @CalendarID = 1,
                      @IsInner = 0,
                      @IgnoreTurn = 1

delete pChangeMSP

insert pChangeMSP 
select m.FileID,
       i.InstitutionID,
       m.INN_UL as INN,
       m.KatMSP,
	   m.DateMSP,
	   '19000101',
	   1,
	   0
from (select r.InstOwnerID, sum(abs(isnull(Rest,0))) as Rest 
          from pRestCalendarOut rco  WITH(NOLOCK)
		  join tResource    r WITH(NOLOCK) on rco.BalResourceID = r.ResourceID
         where spid=@@spid
         group by r.InstOwnerID
        ) x
  join tInstitution i WITH(NOLOCK) on i.InstitutionID = x.InstOwnerID 
  join vabs2.MSP.dbo.tMSP m WITH(NOLOCK) on m.INN_UL = i.INN
  left join tRMSPInstitution ri WITH(NOLOCK) on ri.InstitutionID = i.InstitutionID 
 where x.Rest>0
   and i.PropDealPart = 1
   and isnull(ri.InstitutionID,0) = 0
union
select m.FileID ,
       i.InstitutionID,
       m.INN_IP as INN ,
       m.KatMSP,
	   m.DateMSP,
	   '19000101',
	   1,
	   0

from (select r.InstOwnerID, sum(abs(isnull(Rest,0))) as Rest 
          from pRestCalendarOut rco  WITH(NOLOCK)
		  join tResource    r WITH(NOLOCK) on rco.BalResourceID = r.ResourceID
         where spid=@@spid
         group by r.InstOwnerID
        ) x
  join tInstitution i WITH(NOLOCK) on i.InstitutionID = x.InstOwnerID 
  join vabs2.MSP.dbo.tMSP m WITH(NOLOCK) on m.INN_IP = i.INN
  left join tRMSPInstitution ri WITH(NOLOCK) on ri.InstitutionID = i.InstitutionID 
 where x.Rest>0
 and isnull(ri.InstitutionID,0) = 0
 and i.PropDealPart = 0
 union
select '',
       isnull(ri.InstitutionID,0),
       i.INN as INN ,
       ri.Type,
	   ri.DateStart,
	   ri.DateEnd,
	   2,
	   0
  from (select r.InstOwnerID, sum(abs(isnull(Rest,0))) as Rest 
          from pRestCalendarOut rco  WITH(NOLOCK)
		  join tResource    r WITH(NOLOCK) on rco.BalResourceID = r.ResourceID
         where spid=@@spid
         group by r.InstOwnerID
        ) x
  join tInstitution i WITH(NOLOCK) on i.InstitutionID = x.InstOwnerID 
  join tRMSPInstitution ri WITH(NOLOCK) on ri.InstitutionID = i.InstitutionID  
 where x.Rest>0
   and i.PropDealPart = 1
   and not exists (select 1 
                      from vabs2.MSP.dbo.tMSP m WITH(NOLOCK) where m.INN_UL = i.INN
					  )
 union
select '',
       isnull(ri.InstitutionID,0),
       i.INN as INN ,
       ri.Type,
	   ri.DateStart,
	   ri.DateEnd,
	   2,
	   0
from (select r.InstOwnerID, sum(abs(isnull(Rest,0))) as Rest 
          from pRestCalendarOut rco  WITH(NOLOCK)
		  join tResource    r WITH(NOLOCK) on rco.BalResourceID = r.ResourceID
         where spid=@@spid
         group by r.InstOwnerID
        ) x
  join tInstitution i WITH(NOLOCK) on i.InstitutionID = x.InstOwnerID 
  join tRMSPInstitution ri WITH(NOLOCK) on ri.InstitutionID = i.InstitutionID 
 where x.Rest>0
   and i.PropDealPart = 0
   and not exists (select 1 
                      from vabs2.MSP.dbo.tMSP m WITH(NOLOCK) where m.INN_IP = i.INN
					  )                    
union
select m.FileID,
       isnull(ri.InstitutionID,0),
       m.INN_UL as INN,
       m.KatMSP,
	   m.DateMSP,
	   '19000101',
	   3,
	   0
from (select r.InstOwnerID, sum(abs(isnull(Rest,0))) as Rest 
          from pRestCalendarOut rco  WITH(NOLOCK)
		  join tResource    r WITH(NOLOCK) on rco.BalResourceID = r.ResourceID
         where spid=@@spid
         group by r.InstOwnerID
        ) x
  join tInstitution i WITH(NOLOCK) on i.InstitutionID = x.InstOwnerID 
join vabs2.MSP.dbo.tMSP m WITH(NOLOCK) on m.INN_UL = i.INN
join tRMSPInstitution ri WITH(NOLOCK) on ri.InstitutionID = i.InstitutionID and ri.DateEnd = '19000101'
 where x.Rest>0
   and i.PropDealPart = 1
   and m.KatMSP <> ri.Type
union
select m.FileID,
       isnull(ri.InstitutionID,0),
       m.INN_IP as INN ,
       m.KatMSP,
	   m.DateMSP,
	   '19000101',
	   3,
	   0

from (select r.InstOwnerID, sum(abs(isnull(Rest,0))) as Rest 
          from pRestCalendarOut rco  WITH(NOLOCK)
		  join tResource    r WITH(NOLOCK) on rco.BalResourceID = r.ResourceID
         where spid=@@spid
         group by r.InstOwnerID
        ) x
  join tInstitution i WITH(NOLOCK) on i.InstitutionID = x.InstOwnerID 
  join vabs2.MSP.dbo.tMSP m WITH(NOLOCK) on m.INN_IP = i.INN
  join tRMSPInstitution ri WITH(NOLOCK) on ri.InstitutionID = i.InstitutionID and ri.DateEnd = '19000101'
 where x.Rest>0
   and m.KatMSP <> ri.Type
   and i.PropDealPart = 0

delete pChangeMSP where fl = 2 and isnull(DateEnd,'19000101') <> '19000101'

 --выбывшие, поменявшие категорию МСП
 if exists(select 1 from pChangeMSP where fl in (2,3))
  begin
      update tRMSPInstitution
         set DateEnd = dateadd(dd, -datepart(dd,convert(varchar,getdate(),112))+9, convert(varchar,getdate(),112)), 
             Comment = 'Выбыл' + cast(getdate() as varchar)         
		from tRMSPInstitution WITH (rowlock, updlock INDEX=XPKtRMSPInstitution)
		join pChangeMSP WITH(NOLOCK) on pChangeMSP.InstitutionID = tRMSPInstitution.InstitutionID 
		where pChangeMSP.fl in (2,3)
  end
	   
if exists (select 1 from pChangeMSP where fl in (1,3))
   begin
         
		 
		 
		 
		 select @ID = min(InstitutionID)
		   from pChangeMSP
          where fl in (1,3)    
         
		 while isnull(@ID,0)<>0
		       begin
		             select @Dt = DateStart
					        ,@KatMSP =  KatMSP
							,@Comment = case when fl = 1 then 'Добавлен из файла '+convert(varchar,FileID)+' ' + cast(getdate() as varchar)
							                 when fl = 3 then 'Смена категории ' + cast(getdate() as varchar)+' из файла '+convert(varchar,FileID)    
										end	  
					   from pChangeMSP
					   where InstitutionID = @id
		              
					  
		 
		               exec @RetVal = RMSPInst_Insert  @RMSPInstitutionID = @RMSPInstitutionID out
					                         ,@InstitutionID = @ID
											 ,@DateStart = @dt
											 ,@Type = @KatMSP
											 ,@Comment = @Comment

                       if @RetVal != 0
                          begin
                               update pChangeMSP
							      set RetVal = @RetVal
								where InstitutionID = @ID
								  and fl in (1,3)    
                          end

         
		 
		 select @ID = min(InstitutionID)
		   from pChangeMSP
          where fl in (1,3)    
		    and InstitutionID>@instID
			end
   end
ENDCUR_:

if @RetVal = 0
begin

print '123'

exec [dbo].[SendEmail] 
@e_mail = 'andreeva@besteffortsbank.ru'--uso@besteffortsbank.ru
,@copy_recipients = ''--@copy_recipients
,@subject = 'Реестр субъектов малого и среднего предпринимательства'                                    
,@message = 'Реестр субъектов малого и среднего предпринимательства успешно обновлен'
,@importance = 'Normal'
,@sensitivity = 'Normal'


--delete vabs2.MSP.dbo.tMSP

end
else
begin

print '456'

exec [dbo].[SendEmail] 
@e_mail = 'andreeva@besteffortsbank.ru'--uso@besteffortsbank.ru
,@copy_recipients = ''--@copy_recipients
,@subject = 'Реестр субъектов малого и среднего предпринимательства'                                    
,@message = 'Ошибка при загрузке реестра субъектов малого и среднего предпринимательства, см. pChangeMSP'
,@importance = 'Normal'
,@sensitivity = 'Normal'

end

      
return 0
go
grant all on MSPInst to public
--select * from #Res



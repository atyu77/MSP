alter proc SverkaMSP @dateStart smalldatetime,
        @dateEnd   smalldatetime
as
/*
declare @dateStart smalldatetime,
        @dateEnd   smalldatetime,
		@RetVal    int
	
 select @dateStart = '20220701'
        ,@dateEnd = '20220731'
		,@RetVal = 0
	*/	

declare @RetVal    int
select @RetVal = 0


delete pSverkaMSP 

if exists (select 1 from vabs2.MSP.dbo.pSverkaMSP where substring(ExecDate,7,4)+substring(ExecDate,4,2)+substring(ExecDate,1,2) = @dateEnd)
begin

exec FCD_GL_FindLstPropByType
       @Type= 634  --PROP_TYPE_RMSPINST


insert pSverkaMSP
--нет в реестре МСП в Диасофте
select sv.INN, sv.Kat, sv.DateStart,'не найден в Диасофте в списке МСП или отсутствуют открытые счета в Диасофте' as Descr
  from vabs2.MSP.dbo.pSverkaMSP  sv
  left join 
 (select i.INN
    from tRMSPInstitution  r WITH (NOLOCK)
    join tInstitution      i WITH(NOLOCK) on r.InstitutionID = i.InstitutionID
    join pAPI_Property_Value p WITH (NOLOCK) on p.Value = r.Type 
    join tResource r1 WITH(NOLOCK) on r1.InstOwnerID = r.InstitutionID         
   where 1=1
     and  p.SPID  = @@SPID
     and r1.BalanceID = 2140
     and left(r1.Brief,3) in ('407','408') 
     and (isnull(r1.DateEnd,'19000101')='19000101'                
      or  isnull(r1.DateEnd,'19000101') >=@dateStart 
         )	
     and r.DateStart = (select max(r1.DateStart)
	                      from tRMSPInstitution  r1 WITH (NOLOCK)
						 where r1.InstitutionID = i.InstitutionID
						   and r1.DateStart<=@dateEnd
                        ) 
     and r.DateEnd = '19000101'						
  ) d on d.INN = sv.INN  --
where 1 = 1
  and isnull(d.INN,'') = ''
  and sv.Kat <> 'Не является субъектом МСП'  
                  
union
select distinct sv.INN, sv.Kat, sv.DateStart, 'В Диасофте является МСП, в списке налоговой - нет' as Descr
  from vabs2.MSP.dbo.pSverkaMSP  sv
  join
  (select i.INN, r.DateStart, r.dateEnd, p.Name as Kat
    from tRMSPInstitution  r WITH (NOLOCK)
    join tInstitution      i WITH(NOLOCK) on r.InstitutionID = i.InstitutionID
    join pAPI_Property_Value p WITH (NOLOCK) on p.Value = r.Type 
    join tResource r1 WITH(NOLOCK) on r1.InstOwnerID = r.InstitutionID         
   where 1=1
     and  p.SPID  = @@SPID
     and r1.BalanceID = 2140
     and left(r1.Brief,3) in ('407','408') 
     and (isnull(r1.DateEnd,'19000101')='19000101'                
      or  isnull(r1.DateEnd,'19000101') >=@dateStart
	     )		   
     and r.DateStart = (select max(r1.DateStart)
	                      from tRMSPInstitution  r1 WITH (NOLOCK)
						 where r1.InstitutionID = i.InstitutionID
						   and r1.DateStart<=@dateEnd
     and r.DateEnd = '19000101'
                        ) 
        
   ) d on d.INN = sv.INN  --
 where 1 = 1
   and sv.Kat = 'Не является субъектом МСП'  
 union
 select d.INN, d.Type, d.DateStart,'В Диасофте является МСП, не найден в списке налоговой' as Descr
   from
 (select i.INN, r.DateStart, r.dateEnd, p.Name as Type
    from tRMSPInstitution  r WITH (NOLOCK)
    join tInstitution      i WITH(NOLOCK) on r.InstitutionID = i.InstitutionID
    join pAPI_Property_Value p WITH (NOLOCK) on p.Value = r.Type 
    join tResource r1 WITH(NOLOCK) on r1.InstOwnerID = r.InstitutionID         
   where 1=1
     and  p.SPID  = @@SPID
     and r1.BalanceID = 2140
     and left(r1.Brief,3) in ('407','408') 
     and (isnull(r1.DateEnd,'19000101')='19000101'                
      or  isnull(r1.DateEnd,'19000101') >=@dateStart
	     )		   
     and r.DateStart = (select max(r1.DateStart)
	                      from tRMSPInstitution  r1 WITH (NOLOCK)
						 where r1.InstitutionID = i.InstitutionID
						   and r1.DateStart<=@dateEnd)
     and r.DateEnd = '19000101'
                        ) d
 left join  vabs2.MSP.dbo.pSverkaMSP  sv on sv.INN = d.INN
 where 1 = 1
  and isnull(sv.INN,'') = ''
  --не совпадают категории МСП
  union
select d.INN, d.Type, d.DateStart,'Не совпадают категории МСП' as Descr
   from
 (select i.INN, r.DateStart, r.dateEnd, p.Name as Type 
    from tRMSPInstitution  r WITH (NOLOCK)
    join tInstitution      i WITH(NOLOCK) on r.InstitutionID = i.InstitutionID
    join pAPI_Property_Value p WITH (NOLOCK) on p.Value = r.Type 
    join tResource r1 WITH(NOLOCK) on r1.InstOwnerID = r.InstitutionID         
   where 1=1
     and  p.SPID  = @@SPID
     and r1.BalanceID = 2140
     and left(r1.Brief,3) in ('407','408') 
     and (isnull(r1.DateEnd,'19000101')='19000101'                
      or  isnull(r1.DateEnd,'19000101') >=@dateStart
	     )		   
     and r.DateStart = (select max(r1.DateStart)
	                      from tRMSPInstitution  r1 WITH (NOLOCK)
						 where r1.InstitutionID = i.InstitutionID
						   and r1.DateStart<=@dateEnd)
     --and r.DateEnd = '19000101'
                        ) d
 join  vabs2.MSP.dbo.pSverkaMSP  sv on sv.INN = d.INN and sv.Kat <> d.Type
 where 1 = 1
  and sv.Kat <> 'Не является субъектом МСП' 
 
 end
 else
 insert pSverkaMSP
 select '','','19000101','Не загружен справочник МСП из налоговой'
 return @RetVal
go 
grant all on SverkaMSP to public
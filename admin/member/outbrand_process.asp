<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp"-->  
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/checkAllowIPWithLog.asp" -->
<%
dim yyyymm,mode,makerid

dim refer
refer = request.ServerVariables("HTTP_REFERER")

mode    = requestCheckvar(request.Form("mode"),30)
yyyymm  = requestCheckvar(request.Form("yyyymm"),10)
makerid = requestCheckvar(request.Form("makerid"),32)

dim thismonthday, reqmonthday
thismonthday = dateserial(Left(CStr(now()),4),Mid(CStr(now()),6,2),1)
reqmonthday  = dateserial(Left(yyyymm,4),right(yyyymm,2),1)

if reqmonthday>=thismonthday then
	response.write "<script>alert('생성할 수 없습니다. - 한달 지난 후에 생성 가능합니다.');</script>"
	response.write "<script>history.back();</script>"
	dbget.close()	:	response.End
end if

dim sqlstr
dim paramInfo, retParamInfo, RetErr, retErrStr

if mode="makeoutbrand" then
    '' 반품주문건 제외 (2015/07/20) too long time
    sqlstr = " exec db_partner.[dbo].[sp_Ten_makeOUTBrand] '" + yyyymm + "'"
    dbget.Execute sqlstr
    
    ''분리 2015/12/15
    sqlstr = " exec db_partner.[dbo].[sp_Ten_makeOUTBrand_LastSellDAte] '" + yyyymm + "'"
    dbget.Execute sqlstr
    
elseif mode="prcoutbrand" then
    'response.write "작업중"
    'response.end
    paramInfo = Array(Array("@RETURN_VALUE",adInteger,adParamReturnValue,,0) _
            ,Array("@yyyymm"	,adVarchar, adParamInput,7, yyyymm) _
            ,Array("@makerid"	,adVarchar, adParamInput,32, makerid) _
            ,Array("@retErrStr"	,adVarchar, adParamOutput,300, "") _
    	)
    sqlStr = "db_partner.dbo.sp_Ten_OutBrandProc" 
    retParamInfo = fnExecSPOutput(sqlStr,paramInfo)
       
    RetErr    = GetValue(retParamInfo, "@RETURN_VALUE") ' 에러코드 or IDX
    retErrStr  = GetValue(retParamInfo, "@retErrStr")   ' 에러내용
    
    ''------------------------
    rw "RetErr="&RetErr
    rw "retErrStr="&retErrStr
    
    if (RetErr>0) then ''SCM 로그인 설정 세팅
        paramInfo = Array(Array("@RETURN_VALUE",adInteger,adParamReturnValue,,0) _
            ,Array("@yyyymm"	,adVarchar, adParamInput,7, yyyymm) _
            ,Array("@makerid"	,adVarchar, adParamInput,32, makerid) _
            ,Array("@retErrStr"	,adVarchar, adParamOutput,300, "") _
    	)
        sqlStr = "db_partner.dbo.sp_Ten_OutBrandScmNotUsingProc" 
        retParamInfo = fnExecSPOutput(sqlStr,paramInfo)
           
        RetErr    = GetValue(retParamInfo, "@RETURN_VALUE") ' 에러코드 or IDX
        retErrStr  = GetValue(retParamInfo, "@retErrStr")   ' 에러내용
        
        rw "RetErr="&RetErr
        rw "retErrStr="&retErrStr
    end if
    
    dbget.close() : response.end
elseif mode="prcscmnotusing" then 
    paramInfo = Array(Array("@RETURN_VALUE",adInteger,adParamReturnValue,,0) _
            ,Array("@yyyymm"	,adVarchar, adParamInput,7, yyyymm) _
            ,Array("@makerid"	,adVarchar, adParamInput,32, makerid) _
            ,Array("@retErrStr"	,adVarchar, adParamOutput,300, "") _
    	)
    sqlStr = "db_partner.dbo.sp_Ten_OutBrandScmNotUsingProc" 
    retParamInfo = fnExecSPOutput(sqlStr,paramInfo)
       
    RetErr    = GetValue(retParamInfo, "@RETURN_VALUE") ' 에러코드 or IDX
    retErrStr  = GetValue(retParamInfo, "@retErrStr")   ' 에러내용
    
    ''------------------------
    rw "RetErr="&RetErr
    rw "retErrStr="&retErrStr
    dbget.close() : response.end
elseif mode="makeoutbrand" then '' Not Using
	''삭제
	sqlstr = " delete from [db_partner].[dbo].tbl_outbrand"
	sqlstr = sqlstr + " where yyyymm='" + yyyymm + "'"
	rsget.Open sqlStr,dbget,1

	''신상품갯수
	sqlstr = " insert into [db_partner].[dbo].tbl_outbrand"
	sqlstr = sqlstr + " (yyyymm,makerid,makername,newitemcount)"
	sqlstr = sqlstr + " select '" + yyyymm + "',c.userid, c.socname_kor, IsNULL(T.cnt,0) as regcnt"
	sqlstr = sqlstr + " from [db_user].[dbo].tbl_user_c c"
	sqlstr = sqlstr + " left join ("
	sqlstr = sqlstr + " 	select makerid, count(itemid) as cnt"
	sqlstr = sqlstr + " 	from [db_item].[dbo].tbl_item "
	sqlstr = sqlstr + " 	where datediff(m,regdate,'" + yyyymm + "-01')<3"
	sqlstr = sqlstr + " 	group by makerid"
	sqlstr = sqlstr + " 	) as T on c.userid=T.makerid"
	sqlstr = sqlstr + " where c.userdiv<10"

	rsget.Open sqlStr,dbget,1

	''사용상품갯수
	sqlstr = " update [db_partner].[dbo].tbl_outbrand"
	sqlstr = sqlstr + " set [db_partner].[dbo].tbl_outbrand.usingitemcount = IsNULL(T.usingitemcount,0)"
	sqlstr = sqlstr + " from (select makerid, count(itemid) as usingitemcount"
	sqlstr = sqlstr + " 	from [db_item].[dbo].tbl_item "
	sqlstr = sqlstr + " 	where isusing='Y'"
	sqlstr = sqlstr + " 	group by makerid"
	sqlstr = sqlstr + " ) as T"
	sqlstr = sqlstr + " where [db_partner].[dbo].tbl_outbrand.yyyymm='" + yyyymm + "'"
	sqlstr = sqlstr + " and [db_partner].[dbo].tbl_outbrand.makerid=T.makerid"

	rsget.Open sqlStr,dbget,1

	''온라인 정산금액
	sqlstr = " update [db_partner].[dbo].tbl_outbrand"
	sqlstr = sqlstr + " set [db_partner].[dbo].tbl_outbrand.lastonjungsansum = IsNULL(T.ttljungsan,0)"
	sqlstr = sqlstr + " from (select designerid, "
	sqlstr = sqlstr + " sum(ub_totalsuplycash+me_totalsuplycash +wi_totalsuplycash+sh_totalsuplycash+et_totalsuplycash+dlv_totalsuplycash)"
	sqlstr = sqlstr + " as ttljungsan"
	sqlstr = sqlstr + " from [db_jungsan].[dbo].tbl_designer_jungsan_master"
	sqlstr = sqlstr + " where datediff(m,yyyymm + '-01','" + yyyymm + "-01')<3"
	sqlstr = sqlstr + " and cancelyn='N'"
	sqlstr = sqlstr + " group by designerid"
	sqlstr = sqlstr + " ) as T"
	sqlstr = sqlstr + " where [db_partner].[dbo].tbl_outbrand.yyyymm='" + yyyymm + "'"
	sqlstr = sqlstr + " and [db_partner].[dbo].tbl_outbrand.makerid=T.designerid"

	rsget.Open sqlStr,dbget,1

	''오프라인 정산금액
	''
	''


	''온라인 반품주문건
	sqlstr = " update [db_partner].[dbo].tbl_outbrand"
	sqlstr = sqlstr + " set [db_partner].[dbo].tbl_outbrand.lastminuscnt = IsNULL(T.lastminuscnt,0)"
	sqlstr = sqlstr + " ,[db_partner].[dbo].tbl_outbrand.lastminussum = IsNULL(T.lastminussum,0)"
	sqlstr = sqlstr + " from (select makerid,count(itemno) as lastminuscnt, sum(itemno*itemcost) as lastminussum"
	sqlstr = sqlstr + " 	from "
	sqlstr = sqlstr + " 	[db_order].[dbo].tbl_order_master m,"
	sqlstr = sqlstr + " 	[db_order].[dbo].tbl_order_detail d"
	sqlstr = sqlstr + " 	where m.orderserial=d.orderserial"
	sqlstr = sqlstr + " 	and datediff(m,m.regdate,'" + yyyymm + "-01')<7"
	sqlstr = sqlstr + " 	and d.itemid<>0"
	sqlstr = sqlstr + " 	and m.cancelyn='N'"
	sqlstr = sqlstr + " 	and d.cancelyn<>'Y'"
	sqlstr = sqlstr + " 	and d.itemno<0"
	sqlstr = sqlstr + " 	group by makerid"
	sqlstr = sqlstr + " ) as T"
	sqlstr = sqlstr + " where [db_partner].[dbo].tbl_outbrand.yyyymm='" + yyyymm + "'"
	sqlstr = sqlstr + " and [db_partner].[dbo].tbl_outbrand.makerid=T.makerid"

	rsget.Open sqlStr,dbget,1

    ''최종판매월 //2014/06/18
    dim lastsellBase : lastsellBase = LEFT(CStr(dateAdd("yyyy",-1,yyyymm+"-01")),10)
    
    ''
    sqlstr = " DECLARE @tmpTbl TABLE ("&VbCRLF
    sqlstr = sqlstr + " 	makerid varchar(32)"&VbCRLF
    sqlstr = sqlstr + " 	,targetGbn varchar(10)"&VbCRLF
    sqlstr = sqlstr + " 	,mSellDT varchar(10)"&VbCRLF
    sqlstr = sqlstr + " );"&VbCRLF
    sqlstr = sqlstr + " insert into @tmpTbl"&VbCRLF
    sqlstr = sqlstr + " select makerid,targetGbn,Max(yyyymmdd)  as mSellOF"&VbCRLF
    sqlstr = sqlstr + " from [DBDATAMART].db_datamart.dbo.vw_orderLog_chulgoDate"&VbCRLF
    sqlstr = sqlstr + " where yyyymmdd>='"&lastsellBase&"'"&VbCRLF
    sqlstr = sqlstr + " and targetGbn in ('ON','OF')"&VbCRLF
    sqlstr = sqlstr + " and itemno>0 --and itemid<>0"&VbCRLF
    sqlstr = sqlstr + " and isNULL(makerid,'')<>''"&VbCRLF
    sqlstr = sqlstr + " group by makerid,targetGbn"&VbCRLF
    sqlstr = sqlstr + " ;"&VbCRLF
    sqlstr = sqlstr + " update O"&VbCRLF
    sqlstr = sqlstr + " set lastsellDateON=T.mSellDT"&VbCRLF
    sqlstr = sqlstr + " from [db_partner].[dbo].tbl_outbrand O"&VbCRLF
    sqlstr = sqlstr + "     Join tmpTbl T"&VbCRLF
    sqlstr = sqlstr + "     on O.yyyymm='" + yyyymm + "'"&VbCRLF
    sqlstr = sqlstr + "     and O.makerid=T.makerid"&VbCRLF
    sqlstr = sqlstr + "     and T.targetGbn='ON'"&VbCRLF
    sqlstr = sqlstr + " ;"&VbCRLF
    sqlstr = sqlstr + " update O"&VbCRLF
    sqlstr = sqlstr + " set lastsellDateOF=T.mSellDT"&VbCRLF
    sqlstr = sqlstr + " from [db_partner].[dbo].tbl_outbrand O"&VbCRLF
    sqlstr = sqlstr + "     Join tmpTbl T"&VbCRLF
    sqlstr = sqlstr + "     on O.yyyymm='" + yyyymm + "'"&VbCRLF
    sqlstr = sqlstr + "     and O.makerid=T.makerid"&VbCRLF
    sqlstr = sqlstr + "     and T.targetGbn='OF'"&VbCRLF
    dbget.Execute sqlstr
    
	''레벨설정-5
	sqlstr = " update [db_partner].[dbo].tbl_outbrand"
	sqlstr = sqlstr + " set makerlevel=5"
	sqlstr = sqlstr + " where yyyymm='" + yyyymm + "'"
	sqlstr = sqlstr + " and (lastonjungsansum<500000"
	sqlstr = sqlstr + " or newitemcount<1)"

	rsget.Open sqlStr,dbget,1

	''레벨설정-3
	sqlstr = " update [db_partner].[dbo].tbl_outbrand"
	sqlstr = sqlstr + " set makerlevel=3"
	sqlstr = sqlstr + " where makerid in ("
	sqlstr = sqlstr + " select userid from [db_user].[dbo].tbl_user_c"
	sqlstr = sqlstr + " where datediff(m,regdate,'" + yyyymm + "-01')<4"
	sqlstr = sqlstr + " )"

	rsget.Open sqlStr,dbget,1
end if
%>
<script language="javascript">
alert('저장 되었습니다.');
location.replace('<%= refer %>');
</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->
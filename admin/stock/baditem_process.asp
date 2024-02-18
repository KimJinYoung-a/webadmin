<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->

<%

dim refer
refer = request.ServerVariables("HTTP_REFERER")

dim mode,yyyymmdd,itemgubun,itemid,itemoption
dim sqlStr, i



mode = request("mode")
yyyymmdd    = request("yyyymmdd")
itemgubun   = request("itemgubun")
itemid      = request("itemid")
itemoption  = request("itemoption")

dim erritemExists
dim errcsno, errbaditemno, errrealcheckno, erretcno, toterrno
dim errcode

erritemExists = false
errcsno = 0 
errbaditemno = 0 
errrealcheckno = 0 
erretcno = 0 
toterrno = 0 
errcode = "000"

dim lastitemExists
dim lastyyyymm, stockbaseyyyymmdd
lastitemExists = false

sqlStr = "select convert(varchar(7),dateadd(m,-2,getdate()),21) as LastYYYYMM"
rsget.Open sqlStr,dbget,1
    lastyyyymm = rsget("LastYYYYMM")
rsget.Close

sqlStr = "select convert(varchar(7),dateadd(m,-1,getdate()),21) + '-01' as BaseYYYYMMDD"
rsget.Open sqlStr,dbget,1
    stockbaseyyyymmdd = rsget("BaseYYYYMMDD")
rsget.Close


if (mode="deldetail") then
    '' Insert 고려..?..
    
    sqlStr = "select top 1 * from [db_summary].[dbo].tbl_erritem_daily_summary" + VbCrlf
    sqlStr = sqlStr + " where yyyymmdd='" + yyyymmdd + "'" + VbCrlf
    sqlStr = sqlStr + " and itemgubun='" + itemgubun + "'" + VbCrlf
    sqlStr = sqlStr + " and itemid=" + CStr(itemid) + "" + VbCrlf
    sqlStr = sqlStr + " and itemoption='" + itemoption + "'" + VbCrlf
    
    rsget.Open sqlStr,dbget,1
    if Not rsget.Eof then
        erritemExists = true
        errbaditemno    = rsget("errbaditemno") 
        errrealcheckno  = rsget("errrealcheckno") 
        erretcno        = rsget("erretcno") 
        toterrno        = rsget("toterrno") 
    end if
    rsget.Close
    
    if Not erritemExists then
        response.write "<script>alert('등록된 상품이 존재하지 않습니다.'); window.close();</script>"
        dbget.close()	:	response.End
    end if
    
    On Error Resume Next
    dbget.beginTrans
    
    If Err.Number = 0 Then
        errcode = "001"
    end if
    
    sqlStr = "delete from [db_summary].[dbo].tbl_erritem_daily_summary" + VbCrlf
    sqlStr = sqlStr + " where yyyymmdd='" + yyyymmdd + "'" + VbCrlf
    sqlStr = sqlStr + " and itemgubun='" + itemgubun + "'" + VbCrlf
    sqlStr = sqlStr + " and itemid=" + CStr(itemid) + "" + VbCrlf
    sqlStr = sqlStr + " and itemoption='" + itemoption + "'" + VbCrlf
    
    rsget.Open sqlStr,dbget,1
    
    
    If Err.Number = 0 Then
        errcode = "002"
    end if

    sqlStr = " update [db_summary].[dbo].tbl_daily_logisstock_summary" + VbCrlf
    sqlStr = sqlStr + " set errbaditemno=errbaditemno - (" + CStr(errbaditemno) + ")" + VbCrlf
    sqlStr = sqlStr + " , errrealcheckno=errrealcheckno - (" + CStr(errrealcheckno) + ")" + VbCrlf
    sqlStr = sqlStr + " , erretcno=erretcno - (" + CStr(erretcno) + ")" + VbCrlf
    sqlStr = sqlStr + " , toterrno=toterrno - (" + CStr(toterrno) + ")" + VbCrlf
    sqlStr = sqlStr + " , availsysstock=availsysstock - (" + CStr(errbaditemno + erretcno) + ")" + VbCrlf
    sqlStr = sqlStr + " , realstock=realstock - (" + CStr(toterrno) + ")" + VbCrlf
    
    sqlStr = sqlStr + " where yyyymmdd='" + yyyymmdd + "'" + VbCrlf
    sqlStr = sqlStr + " and itemgubun='" + itemgubun + "'" + VbCrlf
    sqlStr = sqlStr + " and itemid=" + CStr(itemid) + "" + VbCrlf
    sqlStr = sqlStr + " and itemoption='" + itemoption + "'" + VbCrlf
    
    rsget.Open sqlStr,dbget,1
    
    If Err.Number = 0 Then
        errcode = "003"
    end if
    '' update [db_summary].[dbo].tbl_monthly_logisstock_summary
    sqlStr = " update [db_summary].[dbo].tbl_monthly_logisstock_summary" + VbCrlf
    sqlStr = sqlStr + " set errbaditemno=errbaditemno - (" + CStr(errbaditemno) + ")" + VbCrlf
    sqlStr = sqlStr + " , errrealcheckno=errrealcheckno - (" + CStr(errrealcheckno) + ")" + VbCrlf
    sqlStr = sqlStr + " , erretcno=erretcno - (" + CStr(erretcno) + ")" + VbCrlf
    sqlStr = sqlStr + " , toterrno=toterrno - (" + CStr(toterrno) + ")" + VbCrlf
    sqlStr = sqlStr + " , availsysstock=availsysstock - (" + CStr(errbaditemno + erretcno) + ")" + VbCrlf
    sqlStr = sqlStr + " , realstock=realstock - (" + CStr(toterrno) + ")" + VbCrlf
    sqlStr = sqlStr + " where yyyymm='" + Left(yyyymmdd,7) + "'" + VbCrlf
    sqlStr = sqlStr + " and itemgubun='" + itemgubun + "'" + VbCrlf
    sqlStr = sqlStr + " and itemid=" + CStr(itemid) + "" + VbCrlf
    sqlStr = sqlStr + " and itemoption='" + itemoption + "'" + VbCrlf
    
    rsget.Open sqlStr,dbget,1
    
    If Err.Number = 0 Then
        errcode = "004"
    end if
    '' update [db_summary].[dbo].tbl_LAST_monthly_logisstock
    sqlStr = " update [db_summary].[dbo].tbl_LAST_monthly_logisstock" + VbCrlf
    sqlStr = sqlStr + " set errbaditemno=errbaditemno - (" + CStr(errbaditemno) + ")" + VbCrlf
    sqlStr = sqlStr + " , errrealcheckno=errrealcheckno - (" + CStr(errrealcheckno) + ")" + VbCrlf
    sqlStr = sqlStr + " , erretcno=erretcno - (" + CStr(erretcno) + ")" + VbCrlf
    sqlStr = sqlStr + " , toterrno=toterrno - (" + CStr(toterrno) + ")" + VbCrlf
    sqlStr = sqlStr + " , availsysstock=availsysstock - (" + CStr(errbaditemno + erretcno) + ")" + VbCrlf
    sqlStr = sqlStr + " , realstock=realstock - (" + CStr(toterrno) + ")" + VbCrlf
    sqlStr = sqlStr + " where lastyyyymm>='" + Left(yyyymmdd,7) + "'" + VbCrlf
    sqlStr = sqlStr + " and itemgubun='" + itemgubun + "'" + VbCrlf
    sqlStr = sqlStr + " and itemid=" + CStr(itemid) + "" + VbCrlf
    sqlStr = sqlStr + " and itemoption='" + itemoption + "'" + VbCrlf
    
    rsget.Open sqlStr,dbget,1
    
    If Err.Number = 0 Then
        errcode = "005"
    end if
    '' update [db_summary].[dbo].tbl_current_logisstock_summary
    sqlStr = " update [db_summary].[dbo].tbl_current_logisstock_summary" + VbCrlf
    sqlStr = sqlStr + " set errbaditemno=errbaditemno - (" + CStr(errbaditemno) + ")" + VbCrlf
    sqlStr = sqlStr + " , errrealcheckno=errrealcheckno - (" + CStr(errrealcheckno) + ")" + VbCrlf
    sqlStr = sqlStr + " , erretcno=erretcno - (" + CStr(erretcno) + ")" + VbCrlf
    sqlStr = sqlStr + " , toterrno=toterrno - (" + CStr(toterrno) + ")" + VbCrlf
    sqlStr = sqlStr + " , availsysstock=availsysstock - (" + CStr(errbaditemno + erretcno) + ")" + VbCrlf
    sqlStr = sqlStr + " , realstock=realstock - (" + CStr(toterrno) + ")" + VbCrlf
    sqlStr = sqlStr + " where itemgubun='" + itemgubun + "'" + VbCrlf
    sqlStr = sqlStr + " and itemid=" + CStr(itemid) + "" + VbCrlf
    sqlStr = sqlStr + " and itemoption='" + itemoption + "'" + VbCrlf
    
    rsget.Open sqlStr,dbget,1
    
    
    If Err.Number = 0 Then
        dbget.CommitTrans
    Else
        dbget.RollBackTrans
        response.write "<script>alert('데이타를 저장하는 도중에 에러가 발생하였습니다.(에러코드 : " + CStr(errcode) + ")')</script>"
        response.write "<script>window.close()</script>"
        dbget.close()	:	response.End
    End If
    on error Goto 0
elseif (mode="refreshdetail") then
    response.write "관리자 문의 요망 - refreshdetail"
    dbget.close()	:	response.End
    
    sqlStr = " select top 1 * from [db_summary].[dbo].tbl_Last_monthly_logisstock " + VbCrlf
    sqlStr = sqlStr + "     where itemgubun='" + itemgubun + "'" + VbCrlf
    sqlStr = sqlStr + "     and itemid=" + CStr(itemid) + "" + VbCrlf
    sqlStr = sqlStr + "     and itemoption='" + itemoption + "'" + VbCrlf
    
    rsget.Open sqlStr,dbget,1
    if Not rsget.Eof then
        lastitemExists = true
    end if
    rsget.Close
    
 
    On Error Resume Next
    dbget.beginTrans
    
    ''일별 업데이트..추가..
    
    
    If Err.Number = 0 Then
        errcode = "001"
    end if
    
    ''월별 dummy 값 입력
    sqlStr = " insert into [db_summary].[dbo].tbl_monthly_logisstock_summary" + VbCrlf
    sqlStr = sqlStr + " (yyyymm, itemgubun, itemid, itemoption)" + VbCrlf
    sqlStr = sqlStr + " select T.yyyymm, T.itemgubun, T.itemid, T.itemoption" + VbCrlf
    sqlStr = sqlStr + " from (" + VbCrlf
    sqlStr = sqlStr + "     select convert(varchar(7),yyyymmdd) as yyyymm, itemgubun,itemid,itemoption" + VbCrlf
    sqlStr = sqlStr + "     from [db_summary].[dbo].tbl_erritem_daily_summary " + VbCrlf
    sqlStr = sqlStr + "     where itemgubun='" + itemgubun + "'" + VbCrlf
    sqlStr = sqlStr + "     and itemid=" + CStr(itemid) + "" + VbCrlf
    sqlStr = sqlStr + "     and itemoption='" + itemoption + "'" + VbCrlf
    sqlStr = sqlStr + "     group by convert(varchar(7),yyyymmdd),itemgubun,itemid,itemoption " + VbCrlf
    sqlStr = sqlStr + " ) T" + VbCrlf
    sqlStr = sqlStr + " left join [db_summary].[dbo].tbl_monthly_logisstock_summary s" + VbCrlf
    sqlStr = sqlStr + "     on T.yyyymm=s.yyyymm" + VbCrlf
    sqlStr = sqlStr + "     and T.itemgubun=s.itemgubun" + VbCrlf
    sqlStr = sqlStr + "     and T.itemid=s.itemid" + VbCrlf
    sqlStr = sqlStr + "     and T.itemoption=s.itemoption" + VbCrlf
    sqlStr = sqlStr + " where s.yyyymm Is NULL" 
    
    rsget.Open sqlStr,dbget,1
    
    
    If Err.Number = 0 Then
        errcode = "002"
    end if
    
    ''월별 에러로그 저장
    '' 값이 없는것 또는 delete 된것 처리위해 0으로 세팅
    sqlStr = " update [db_summary].[dbo].tbl_monthly_logisstock_summary" + VbCrlf
    sqlStr = sqlStr + " set errbaditemno=0" + VbCrlf
    sqlStr = sqlStr + " , errrealcheckno=0" + VbCrlf
    sqlStr = sqlStr + " , erretcno=0" + VbCrlf
    sqlStr = sqlStr + " , toterrno=0" + VbCrlf
    sqlStr = sqlStr + " , availsysstock=totsysstock+0" + VbCrlf
    sqlStr = sqlStr + " , realstock=totsysstock+0" + VbCrlf
    sqlStr = sqlStr + " , lastupdate=getdate()" + VbCrlf
    sqlStr = sqlStr + " where yyyymm<='" + CStr(lastyyyymm) + "'" + VbCrlf
    sqlStr = sqlStr + " and itemgubun='" + itemgubun + "'" + VbCrlf
    sqlStr = sqlStr + " and itemid=" + CStr(itemid) + "" + VbCrlf
    sqlStr = sqlStr + " and itemoption='" + itemoption + "'" + VbCrlf
    
    rsget.Open sqlStr,dbget,1
    
    
    
    If Err.Number = 0 Then
        errcode = "003"
    end if
    
    sqlStr = " update [db_summary].[dbo].tbl_monthly_logisstock_summary" + VbCrlf
    sqlStr = sqlStr + " set errbaditemno=IsNULL(T.errbaditemno,0)" + VbCrlf
    sqlStr = sqlStr + " , errrealcheckno=IsNULL(T.errrealcheckno,0)" + VbCrlf
    sqlStr = sqlStr + " , erretcno=IsNULL(T.erretcno,0)" + VbCrlf
    sqlStr = sqlStr + " , toterrno=IsNULL(T.toterrno,0)" + VbCrlf
    sqlStr = sqlStr + " , availsysstock=totsysstock+IsNULL(T.errbaditemno,0)" + VbCrlf
    sqlStr = sqlStr + " , realstock=totsysstock+IsNULL(T.toterrno,0)" + VbCrlf
    sqlStr = sqlStr + " , lastupdate=getdate()" + VbCrlf
    sqlStr = sqlStr + " from ("
    sqlStr = sqlStr + "         select convert(varchar(7),yyyymmdd) as yyyymm, itemgubun,itemid,itemoption, " + VbCrlf
    sqlStr = sqlStr + "         sum(errbaditemno) as errbaditemno, " + VbCrlf
    sqlStr = sqlStr + "         sum(errrealcheckno) as errrealcheckno, " + VbCrlf
    sqlStr = sqlStr + "         sum(erretcno) as erretcno, " + VbCrlf
    sqlStr = sqlStr + "         sum(toterrno) as toterrno " + VbCrlf
    sqlStr = sqlStr + "         from [db_summary].[dbo].tbl_erritem_daily_summary " + VbCrlf
    sqlStr = sqlStr + "         where convert(varchar(7),yyyymmdd)<='" + CStr(lastyyyymm) + "'" + VbCrlf
    sqlStr = sqlStr + "         and itemgubun='" + itemgubun + "'" + VbCrlf
    sqlStr = sqlStr + "         and itemid=" + CStr(itemid) + "" + VbCrlf
    sqlStr = sqlStr + "         and itemoption='" + itemoption + "'" + VbCrlf
    sqlStr = sqlStr + "         group by convert(varchar(7),yyyymmdd),itemgubun,itemid,itemoption " + VbCrlf
    sqlStr = sqlStr + " ) T " + VbCrlf
    sqlStr = sqlStr + " where [db_summary].[dbo].tbl_monthly_logisstock_summary.yyyymm=T.yyyymm" + VbCrlf
    sqlStr = sqlStr + " and [db_summary].[dbo].tbl_monthly_logisstock_summary.itemgubun=T.itemgubun" + VbCrlf
    sqlStr = sqlStr + " and [db_summary].[dbo].tbl_monthly_logisstock_summary.itemid=T.itemid" + VbCrlf
    sqlStr = sqlStr + " and [db_summary].[dbo].tbl_monthly_logisstock_summary.itemoption=T.itemoption" + VbCrlf
    
    rsget.Open sqlStr,dbget,1
    
    
    If Err.Number = 0 Then
        errcode = "004"
    end if
    
    '' Last_monthly_logisstock Update
    if Not (lastitemExists) then
        sqlStr = " insert into [db_summary].[dbo].tbl_last_monthly_logisstock" + VbCrlf
        sqlStr = sqlStr + " (lastyyyymm, itemgubun, itemid, itemoption)" + VbCrlf
        sqlStr = sqlStr + " values(" + VbCrlf
        sqlStr = sqlStr + " '" + lastyyyymm + "'" + VbCrlf
        sqlStr = sqlStr + ",'" + itemgubun + "'" + VbCrlf
        sqlStr = sqlStr + "," + CStr(itemid) + "" + VbCrlf
        sqlStr = sqlStr + ",'" + itemoption + "'" + VbCrlf
        sqlStr = sqlStr + ")" + VbCrlf
        
        rsget.Open sqlStr,dbget,1
    end if  
    
    If Err.Number = 0 Then
        errcode = "005"
    end if
    
    sqlStr = " update [db_summary].[dbo].tbl_LAST_monthly_logisstock" + VbCrlf
    sqlStr = sqlStr + " set errbaditemno=IsNULL(T.errbaditemno,0)" + VbCrlf
    sqlStr = sqlStr + " ,errrealcheckno=IsNULL(T.errrealcheckno,0)" + VbCrlf
    sqlStr = sqlStr + " ,erretcno=IsNULL(T.erretcno,0)" + VbCrlf
    sqlStr = sqlStr + " ,toterrno=IsNULL(T.toterrno,0)" + VbCrlf
    sqlStr = sqlStr + " ,availsysstock=totsysstock+IsNULL(T.errbaditemno,0)" + VbCrlf
    sqlStr = sqlStr + " ,realstock=totsysstock+IsNULL(T.toterrno,0)" + VbCrlf
    sqlStr = sqlStr + " ,lastupdate=getdate()" + VbCrlf
    sqlStr = sqlStr + " from (" + VbCrlf
    sqlStr = sqlStr + "     select itemgubun, itemid, itemoption " + VbCrlf
    sqlStr = sqlStr + "     ,sum(errbaditemno) as errbaditemno " + VbCrlf
    sqlStr = sqlStr + "     ,sum(errrealcheckno) as errrealcheckno " + VbCrlf
    sqlStr = sqlStr + "     ,sum(erretcno) as erretcno " + VbCrlf
    sqlStr = sqlStr + "     ,sum(toterrno) as toterrno " + VbCrlf
    sqlStr = sqlStr + "     from [db_summary].[dbo].tbl_monthly_logisstock_summary " + VbCrlf
    sqlStr = sqlStr + "     where yyyymm<='" + lastyyyymm + "'" + VbCrlf
    sqlStr = sqlStr + "     and itemgubun='" + itemgubun + "'" + VbCrlf
    sqlStr = sqlStr + "     and itemid=" + CStr(itemid) + "" + VbCrlf
    sqlStr = sqlStr + "     and itemoption='" + itemoption + "'" + VbCrlf
    sqlStr = sqlStr + "     group by itemgubun, itemid, itemoption" 
    sqlStr = sqlStr + "     ) as T"
    sqlStr = sqlStr + " where [db_summary].[dbo].tbl_LAST_monthly_logisstock.itemgubun=T.itemgubun" + VbCrlf
    sqlStr = sqlStr + " and [db_summary].[dbo].tbl_LAST_monthly_logisstock.itemid=T.itemid" + VbCrlf
    sqlStr = sqlStr + " and [db_summary].[dbo].tbl_LAST_monthly_logisstock.itemoption=T.itemoption" + VbCrlf
    
    
    rsget.Open sqlStr,dbget,1

    
    '' update Current_logisstock
    If Err.Number = 0 Then
        errcode = "006"
    end if
    
    sqlStr = " insert into [db_summary].[dbo].tbl_current_logisstock_summary" + VbCrlf
    sqlStr = sqlStr + " (itemgubun, itemid, itemoption)" + VbCrlf
    sqlStr = sqlStr + " select L.itemgubun, L.itemid, L.itemoption " + VbCrlf
    sqlStr = sqlStr + " from [db_summary].[dbo].tbl_LAST_monthly_logisstock L " + VbCrlf
    sqlStr = sqlStr + " left join [db_summary].[dbo].tbl_current_logisstock_summary C " + VbCrlf
    sqlStr = sqlStr + " on L.itemgubun=C.itemgubun"
    sqlStr = sqlStr + " and L.itemid=C.itemid"
    sqlStr = sqlStr + " and L.itemoption=C.itemoption"
    sqlStr = sqlStr + " where C.itemgubun is NULL"
    
    rsget.Open sqlStr,dbget,1
    
    
    If Err.Number = 0 Then
        errcode = "007"
    end if
    
    sqlStr = " update [db_summary].[dbo].tbl_current_logisstock_summary" + VbCrlf
    sqlStr = sqlStr + " set errbaditemno=IsNULL(T.errbaditemno,0)" + VbCrlf
    sqlStr = sqlStr + " ,errrealcheckno=IsNULL(T.errrealcheckno,0)" + VbCrlf
    sqlStr = sqlStr + " ,erretcno=IsNULL(T.erretcno,0)" + VbCrlf
    sqlStr = sqlStr + " ,toterrno=IsNULL(T.toterrno,0)" + VbCrlf
    sqlStr = sqlStr + " ,availsysstock=totsysstock+IsNULL(T.errbaditemno,0)" + VbCrlf
    sqlStr = sqlStr + " ,realstock=totsysstock+IsNULL(T.toterrno,0)" + VbCrlf
    sqlStr = sqlStr + " ,lastupdate=getdate()" + VbCrlf
    sqlStr = sqlStr + " from ("
    sqlStr = sqlStr + "     select L.itemgubun, L.itemid, L.itemoption " + VbCrlf
    sqlStr = sqlStr + "     ,L.errcsno + IsNULL(sum(d.errcsno),0) as errcsno " + VbCrlf
    sqlStr = sqlStr + "     ,L.errbaditemno + IsNULL(sum(d.errbaditemno),0) as errbaditemno " + VbCrlf
    sqlStr = sqlStr + "     ,L.errrealcheckno + IsNULL(sum(d.errrealcheckno),0) as errrealcheckno " + VbCrlf
    sqlStr = sqlStr + "     ,L.erretcno + IsNULL(sum(d.erretcno),0) as erretcno " + VbCrlf
    sqlStr = sqlStr + "     ,L.toterrno + IsNULL(sum(d.toterrno),0) as toterrno " + VbCrlf
    sqlStr = sqlStr + "     from [db_summary].[dbo].tbl_LAST_monthly_logisstock L " + VbCrlf
    sqlStr = sqlStr + "         left join [db_summary].[dbo].tbl_daily_logisstock_summary d" + VbCrlf
    sqlStr = sqlStr + "         on L.itemgubun=d.itemgubun" + VbCrlf
    sqlStr = sqlStr + "         and L.itemid=d.itemid" + VbCrlf
    sqlStr = sqlStr + "         and L.itemoption=d.itemoption" + VbCrlf
    sqlStr = sqlStr + "         and d.yyyymmdd>='" + stockbaseyyyymmdd + "'" + VbCrlf
    sqlStr = sqlStr + "     where L.itemgubun='" + itemgubun + "'" + VbCrlf
    sqlStr = sqlStr + "     and L.itemid=" + Cstr(itemid) + "" + VbCrlf
    sqlStr = sqlStr + "     and L.itemoption='" + itemoption + "'" + VbCrlf
    sqlStr = sqlStr + "     group by L.itemgubun, L.itemid, L.itemoption, L.errcsno, L.errbaditemno, L.errrealcheckno, L.erretcno, L.toterrno" + VbCrlf
    sqlStr = sqlStr + "    ) as T " + VbCrlf
    sqlStr = sqlStr + " where [db_summary].[dbo].tbl_current_logisstock_summary.itemgubun=T.itemgubun" + VbCrlf
    sqlStr = sqlStr + " and [db_summary].[dbo].tbl_current_logisstock_summary.itemid=T.itemid" + VbCrlf
    sqlStr = sqlStr + " and [db_summary].[dbo].tbl_current_logisstock_summary.itemoption=T.itemoption" + VbCrlf
    
    rsget.Open sqlStr,dbget,1
    
    
    
    If Err.Number = 0 Then
        dbget.CommitTrans
    Else
        dbget.RollBackTrans
        response.write "<script>alert('데이타를 저장하는 도중에 에러가 발생하였습니다.(에러코드 : " + CStr(errcode) + ")')</script>"
        response.write "<script>alert('" + Replace(CStr(Err.Description),"'","\'") + "')</script>"
        response.write "<script>window.close()</script>"
        dbget.close()	:	response.End
    End If
    on error Goto 0
else
    response.write "<script>alert('Not Valid mode Key'); window.close();</script>"
    dbget.close()	:	response.End
end if

%>

<script language="javascript">
alert('ok');
opener.location.reload();
window.close();
</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->

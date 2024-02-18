<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/designer/incSessiondesigner.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->

<%

dim refer
refer = request.ServerVariables("HTTP_REFERER")

dim mode,sqlStr
dim statecd, masteridx
dim ipgodate, scheduleipgodate, replycomment
dim beasongdate, songjangdiv, songjangno, songjangname

mode = requestCheckVar(request("mode"),32)
masteridx = requestCheckVar(request("masteridx"),10)
scheduleipgodate = request("scheduleipgodate")
replycomment = html2db(request("replycomment"))
statecd         = requestCheckVar(request("statecd"),10)
beasongdate = requestCheckVar(request("beasongdate"),32)
songjangdiv = requestCheckVar(request("songjangdiv"),32)
songjangname = html2db(request("songjangname"))
songjangno = requestCheckVar(request("songjangno"),60)

dim realitemno, comment
dim itemgubun, itemid, itemoption
realitemno = requestCheckVar(request("realitemno"),10)
comment = html2db(request("comment"))
itemgubun   = requestCheckVar(request("itemgubun"),2)
itemid      = requestCheckVar(request("itemid"),10)
itemoption  = requestCheckVar(request("itemoption"),4)

dim dtstat,detailidx
detailidx = requestCheckVar(request("detailidx"),10)
dtstat=requestCheckVar(request("dtstat"),32)
'reipgodate=request("reipgodate")


dim oldrealitemno

if (mode="modistate") then
	sqlStr = " update [db_storage].[dbo].tbl_ordersheet_master" + vbCrlf
	sqlStr = sqlStr + " set replycomment='" + replycomment + "'"  + vbCrlf
	if (scheduleipgodate<>"") then
    	sqlStr = sqlStr + " ,scheduleipgodate='" + scheduleipgodate + "'" + vbCrlf
    end if
	sqlStr = sqlStr + " ,statecd='" + statecd + "'" + vbCrlf
	if beasongdate<>"" then
		sqlStr = sqlStr + " ,beasongdate='" + beasongdate + "'" + vbCrlf
	end if

	if songjangdiv<>"" then
		sqlStr = sqlStr + " ,songjangdiv='" + songjangdiv + "'" + vbCrlf
	end if

 	if songjangno<>"" then
		sqlStr = sqlStr + " ,songjangno='" + songjangno + "'" + vbCrlf
	end if

    ''??
 	if songjangname<>"" and songjangname<>"선택" then
		sqlStr = sqlStr + " ,songjangname='" + songjangname + "'" + vbCrlf
	end if


	sqlStr = sqlStr + " where idx=" + CStr(masteridx)

	dbget.Execute sqlStr
elseif mode="modidetail" then
    '' 기존 수량
	sqlStr = "select realitemno from [db_storage].[dbo].tbl_ordersheet_detail"
	sqlStr = sqlStr + " where masteridx=" + CStr(masteridx)
	sqlStr = sqlStr + " and itemgubun='" & itemgubun & "'" + vbCrlf
	sqlStr = sqlStr + " and itemid=" & itemid & vbCrlf
	sqlStr = sqlStr + " and itemoption='" & itemoption & "'"

	rsget.Open sqlStr, dbget, 1
	if Not rsget.Eof then
		oldrealitemno = rsget("realitemno")
    end if
	rsget.close


	sqlStr = " update [db_storage].[dbo].tbl_ordersheet_detail" & vbCrlf
	sqlStr = sqlStr + " set realitemno = " & realitemno  & vbCrlf
	sqlStr = sqlStr + " ,comment = '" & comment & "'" & vbCrlf
	sqlStr = sqlStr + " where masteridx=" + CStr(masteridx) + vbCrlf
	sqlStr = sqlStr + " and itemgubun='" & itemgubun & "'" + vbCrlf
	sqlStr = sqlStr + " and itemid=" & itemid & vbCrlf
	sqlStr = sqlStr + " and itemoption='" & itemoption & "'"

	dbget.Execute sqlStr

	sqlStr = " update [db_storage].[dbo].tbl_ordersheet_master" + vbCrlf
	sqlStr = sqlStr + " set jumunsellcash=IsNull(T.totsell,0)" + vbCrlf
	sqlStr = sqlStr + " ,jumunsuplycash=IsNull(T.totsupp,0)" + vbCrlf
	sqlStr = sqlStr + " ,jumunbuycash=IsNull(T.totbuy,0)" + vbCrlf
	sqlStr = sqlStr + " ,totalsellcash=IsNull(T.realtotsell,0)" + vbCrlf
	sqlStr = sqlStr + " ,totalsuplycash=IsNull(T.realtotsupp,0)" + vbCrlf
	sqlStr = sqlStr + " ,totalbuycash=IsNull(T.realtotbuy,0)" + vbCrlf
	sqlStr = sqlStr + " from (select sum(sellcash*baljuitemno) as totsell, " + vbCrlf
	sqlStr = sqlStr + " sum(suplycash*baljuitemno) as totsupp, " + vbCrlf
	sqlStr = sqlStr + " sum(buycash*baljuitemno) as totbuy, " + vbCrlf
	sqlStr = sqlStr + " sum(sellcash*realitemno) as realtotsell, " + vbCrlf
	sqlStr = sqlStr + " sum(suplycash*realitemno) as realtotsupp, " + vbCrlf
	sqlStr = sqlStr + " sum(buycash*realitemno) as realtotbuy " + vbCrlf
	sqlStr = sqlStr + " from [db_storage].[dbo].tbl_ordersheet_detail" + vbCrlf
	sqlStr = sqlStr + " where masteridx="  + CStr(masteridx) + vbCrlf
	sqlStr = sqlStr + " and deldt is null" + vbCrlf
	sqlStr = sqlStr + " ) as T" + vbCrlf
	sqlStr = sqlStr + " where [db_storage].[dbo].tbl_ordersheet_master.idx=" + CStr(masteridx)

	dbget.Execute sqlStr

    ''기주문수량업데이트
    sqlStr = " update [db_summary].[dbo].tbl_current_logisstock_summary" & VbCrlf
    sqlStr = sqlStr & " set preordernofix=preordernofix + " & realitemno-oldrealitemno & VbCrlf
    sqlStr = sqlStr & " where itemgubun='" & itemgubun & "'" & VbCrlf
    sqlStr = sqlStr & " and itemid=" & itemid & VbCrlf
    sqlStr = sqlStr & " and itemoption='" & itemoption & "'" & VbCrlf
    dbget.Execute sqlStr

    dim detail_status,cnt

    sqlStr = " select count(*) as cnt from [db_storage].[dbo].tbl_ordersheet_detail_log where detail_idx='" & CStr(detailidx) & "'"


    rsget.open sqlStr,dbget,1

    if not rsget.eof then
    	cnt = rsget("cnt")
    end if
    rsget.close


    if dtstat="" then
    	detail_status=""

	elseif dtstat="ipt" then
		detail_status= "직접입력"

	elseif dtstat="so" then
		detail_status ="단종"

	elseif dtstat="sso" then
		detail_status ="일시품절"

	end if

    if cnt>0 then
    	sqlStr =" update [db_storage].[dbo].tbl_ordersheet_detail_log " &vbCRLF
    	sqlStr = sqlStr&" set detail_status='" & detail_status & "'" &vbCRLF
    	sqlStr = sqlStr&" ,detail_description ='" & comment & "'" &vbCRLF
    	sqlStr = sqlStr&" where detail_idx='" & CStr(detailidx) & "'"
    else
    	sqlStr =" insert into db_storage.dbo.tbl_ordersheet_detail_log(detail_idx,detail_status,detail_description) " & vbCRLF
    	sqlStr = sqlStr&" values('" & CStr(detailidx) & "','" & detail_status & "','" & comment & "') "
    end if

    	dbget.execute(sqlStr)

    IF dtstat="sso" and isDate(comment) Then
    	sqlStr = "exec [db_storage].[dbo].sp_Ten_StockReipgoSetting '10',"& itemid & ",'"& itemoption &"','"& comment &"'"
    	dbget.Execute(sqlStr)
    End IF
else
    response.write "정의 되지 않았습니다." & mode
    dbget.close()	:	response.End
end if
%>

<script language="javascript">
alert('수정 되었습니다.');
location.replace('<%= refer %>');
</script>

<!-- #include virtual="/lib/db/dbclose.asp" -->

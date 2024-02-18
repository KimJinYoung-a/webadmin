<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description : 재고
' History : 이상구 생성
'			2017.04.12 한용민 수정(보안관련처리)
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<%
dim idx, designer
dim shopid, itemgubunarr
dim shopitemarr, itemoptionarr,realjeagoarr
dim yyyy1,mm1,dd1
dim hh1,nn1,ss1
dim jeagodate
yyyy1 = requestCheckVar(request("yyyy1"),4)
mm1 = requestCheckVar(request("mm1"),2)
dd1 = requestCheckVar(request("dd1"),2)
hh1 = requestCheckVar(request("hh1"),2)
nn1 = requestCheckVar(request("nn1"),2)
ss1 = requestCheckVar(request("ss1"),2)
jeagodate = yyyy1 + "-" + mm1 + "-" + dd1 + " " + hh1 + ":" + nn1 + ":" + ss1

idx      = requestCheckVar(request("idx"),10)
designer = requestCheckVar(request("designer"),32)
shopid	 = requestCheckVar(request("shopid"),32)
itemgubunarr	= request("itemgubunarr")
shopitemarr		= request("shopitemarr")
itemoptionarr	= request("itemoptionarr")
realjeagoarr 	= request("realjeagoarr")

dim i,cnt,sqlStr
dim oldjeago
itemgubunarr = Left(itemgubunarr,Len(itemgubunarr)-1)
shopitemarr = Left(shopitemarr,Len(shopitemarr)-1)
itemoptionarr = Left(itemoptionarr,Len(itemoptionarr)-1)
realjeagoarr = Left(realjeagoarr,Len(realjeagoarr)-1)

itemgubunarr = split(itemgubunarr,"|")
shopitemarr = split(shopitemarr,"|")
itemoptionarr = split(itemoptionarr,"|")
realjeagoarr = split(realjeagoarr,"|")

cnt = ubound(shopitemarr)

if idx="" then
	sqlStr = "select * from [db_shop].[dbo].tbl_shop_realjaego_master where 1=0"
	rsget.Open sqlStr,dbget,1,3
	rsget.AddNew

	rsget("shopid") = shopid
	rsget("makerid") = designer
	rsget("jeagodate") = jeagodate
	rsget("reguserid") = session("ssBctId")

	rsget.update
		idx = rsget("idx")
	rsget.close

	for i=0 to cnt
		sqlStr = "insert into [db_shop].[dbo].tbl_shop_realjaego_detail"
		sqlStr = sqlStr + " (masteridx,makerid,itemgubun,shopitemid,itemoption,realjeago)"
		sqlStr = sqlStr + " values(" +  CStr(idx) + ","
		sqlStr = sqlStr + " '" + designer + "',"
		sqlStr = sqlStr + " '" + requestCheckVar(itemgubunarr(i),2) + "',"
		sqlStr = sqlStr + " " + requestCheckVar(shopitemarr(i),10) + ","
		sqlStr = sqlStr + " '" + requestCheckVar(itemoptionarr(i),4) + "',"
		sqlStr = sqlStr + " " + requestCheckVar(realjeagoarr(i),10) + ")"

		rsget.Open sqlStr,dbget,1

		sqlStr = " update [db_shop].[dbo].tbl_shop_day_stock" + VbCrlF
		sqlStr = sqlStr + " set lastrealdate='" + jeagodate + "'"  + VbCrlF
		sqlStr = sqlStr + " ,lastrealno=" + requestCheckVar(realjeagoarr(i),10)  + VbCrlF
		sqlStr = sqlStr + " ,ipno=0"  + VbCrlF
		sqlStr = sqlStr + " ,reno=0"  + VbCrlF
		sqlStr = sqlStr + " ,upcheipno=0"  + VbCrlF
		sqlStr = sqlStr + " ,upchereno=0"  + VbCrlF
		sqlStr = sqlStr + " ,sellno=0"  + VbCrlF
		sqlStr = sqlStr + " ,currno=0"  + VbCrlF
		sqlStr = sqlStr + " ,sell7days=0"  + VbCrlF
		sqlStr = sqlStr + " ,preorderno=0"  + VbCrlF
		sqlStr = sqlStr + " ,requireno=0"  + VbCrlF
		sqlStr = sqlStr + " ,shortageno=0"  + VbCrlF
		sqlStr = sqlStr + " ,maxsellday=0"  + VbCrlF

		sqlStr = sqlStr + " where shopid='" + shopid + "'"  + VbCrlF
		sqlStr = sqlStr + " and itemgubun='" + requestCheckVar(itemgubunarr(i),2) + "'"  + VbCrlF
		sqlStr = sqlStr + " and itemid=" + requestCheckVar(shopitemarr(i),10)  + VbCrlF
		sqlStr = sqlStr + " and itemoption=" + requestCheckVar(itemoptionarr(i),4)  + VbCrlF

		rsget.Open sqlStr,dbget,1
	next
else
	sqlStr = "update [db_shop].[dbo].tbl_shop_realjaego_master"
	sqlStr = sqlStr + " set jeagodate='" + jeagodate + "'"
	sqlStr = sqlStr + " ,edituserid='" + session("ssBctId") + "'"
	sqlStr = sqlStr + " where idx=" + idx
	rsget.Open sqlStr,dbget,1

	for i=0 to cnt
		oldjeago = -999
		sqlStr = "select top 1 realjeago from [db_shop].[dbo].tbl_shop_realjaego_detail"
		sqlStr = sqlStr + " where masteridx=" + idx
		sqlStr = sqlStr + " and itemgubun='" + requestCheckVar(itemgubunarr(i),2) + "'"
		sqlStr = sqlStr + " and shopitemid=" + requestCheckVar(shopitemarr(i),10)
		sqlStr = sqlStr + " and itemoption='" + requestCheckVar(itemoptionarr(i),4) + "'"
		rsget.Open sqlStr,dbget,1
			if Not rsget.Eof then
				oldjeago = rsget("realjeago")
			end if
		rsget.close

		if CStr(oldjeago)<>realjeagoarr(i) then
			sqlStr = " update [db_shop].[dbo].tbl_shop_realjaego_detail"
			sqlStr = sqlStr + " set realjeago=" + requestCheckVar(realjeagoarr(i),10)
			sqlStr = sqlStr + " where masteridx=" + idx
			sqlStr = sqlStr + " and itemgubun='" + requestCheckVar(itemgubunarr(i),2) + "'"
			sqlStr = sqlStr + " and shopitemid=" + requestCheckVar(shopitemarr(i),10)
			sqlStr = sqlStr + " and itemoption='" + requestCheckVar(itemoptionarr(i),4) + "'"

			rsget.Open sqlStr,dbget,1
		end if

		sqlStr = " update [db_shop].[dbo].tbl_shop_day_stock" + VbCrlF
		sqlStr = sqlStr + " set lastrealdate='" + jeagodate + "'"  + VbCrlF
		sqlStr = sqlStr + " ,lastrealno=" + requestCheckVar(realjeagoarr(i),10)  + VbCrlF
		sqlStr = sqlStr + " ,ipno=0"  + VbCrlF
		sqlStr = sqlStr + " ,reno=0"  + VbCrlF
		sqlStr = sqlStr + " ,upcheipno=0"  + VbCrlF
		sqlStr = sqlStr + " ,upchereno=0"  + VbCrlF
		sqlStr = sqlStr + " ,sellno=0"  + VbCrlF
		sqlStr = sqlStr + " ,currno=0"  + VbCrlF
		sqlStr = sqlStr + " ,sell7days=0"  + VbCrlF
		sqlStr = sqlStr + " ,preorderno=0"  + VbCrlF
		sqlStr = sqlStr + " ,requireno=0"  + VbCrlF
		sqlStr = sqlStr + " ,shortageno=0"  + VbCrlF
		sqlStr = sqlStr + " ,maxsellday=0"  + VbCrlF

		sqlStr = sqlStr + " where shopid='" + shopid + "'"  + VbCrlF
		sqlStr = sqlStr + " and itemgubun='" + requestCheckVar(itemgubunarr(i),2) + "'"  + VbCrlF
		sqlStr = sqlStr + " and itemid=" + requestCheckVar(shopitemarr(i),10)  + VbCrlF
		sqlStr = sqlStr + " and itemoption='" + requestCheckVar(itemoptionarr(i),4) + "'"  + VbCrlF

		rsget.Open sqlStr,dbget,1
	next
end if
%>

<script type='text/javascript'>
	alert('저장 되었습니다.');
	location.replace('brandjaegolist.asp?shopid=<%= shopid %>&makerid=<%= designer %>');
</script>

<!-- #include virtual="/lib/db/dbclose.asp" -->
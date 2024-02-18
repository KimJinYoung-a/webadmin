<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<%
dim idx, designer
dim shopid, itemgubunarr
dim shopitemarr, itemoptionarr,realjeagoarr
dim yyyy1,mm1,dd1
dim hh1,nn1,ss1
dim jeagodate

yyyy1 = request("yyyy1")
mm1 = request("mm1")
dd1 = request("dd1")
hh1 = request("hh1")
nn1 = request("nn1")
ss1 = request("ss1")
jeagodate = yyyy1 + "-" + mm1 + "-" + dd1 + " " + hh1 + ":" + nn1 + ":" + ss1

idx      = request("idx")
designer = request("designer")
shopid	 = request("shopid")
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
		sqlStr = sqlStr + " '" + itemgubunarr(i) + "',"
		sqlStr = sqlStr + " " + shopitemarr(i) + ","
		sqlStr = sqlStr + " '" + itemoptionarr(i) + "',"
		sqlStr = sqlStr + " " + realjeagoarr(i) + ")"

		rsget.Open sqlStr,dbget,1

		sqlStr = " update [db_shop].[dbo].tbl_shop_day_stock" + VbCrlF
		sqlStr = sqlStr + " set lastrealdate='" + jeagodate + "'"  + VbCrlF
		sqlStr = sqlStr + " ,lastrealno=" + realjeagoarr(i)  + VbCrlF
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
		sqlStr = sqlStr + " and itemgubun='" + itemgubunarr(i) + "'"  + VbCrlF
		sqlStr = sqlStr + " and itemid=" + shopitemarr(i)  + VbCrlF
		sqlStr = sqlStr + " and itemoption=" + itemoptionarr(i)  + VbCrlF

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
		sqlStr = sqlStr + " and itemgubun='" + itemgubunarr(i) + "'"
		sqlStr = sqlStr + " and shopitemid=" + shopitemarr(i)
		sqlStr = sqlStr + " and itemoption=" + itemoptionarr(i)
		rsget.Open sqlStr,dbget,1
			if Not rsget.Eof then
				oldjeago = rsget("realjeago")
			end if
		rsget.close

		if CStr(oldjeago)<>realjeagoarr(i) then
			sqlStr = " update [db_shop].[dbo].tbl_shop_realjaego_detail"
			sqlStr = sqlStr + " set realjeago=" + realjeagoarr(i)
			sqlStr = sqlStr + " where masteridx=" + idx
			sqlStr = sqlStr + " and itemgubun='" + itemgubunarr(i) + "'"
			sqlStr = sqlStr + " and shopitemid=" + shopitemarr(i)
			sqlStr = sqlStr + " and itemoption=" + itemoptionarr(i)

			rsget.Open sqlStr,dbget,1
		end if

		sqlStr = " update [db_shop].[dbo].tbl_shop_day_stock" + VbCrlF
		sqlStr = sqlStr + " set lastrealdate='" + jeagodate + "'"  + VbCrlF
		sqlStr = sqlStr + " ,lastrealno=" + realjeagoarr(i)  + VbCrlF
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
		sqlStr = sqlStr + " and itemgubun='" + itemgubunarr(i) + "'"  + VbCrlF
		sqlStr = sqlStr + " and itemid=" + shopitemarr(i)  + VbCrlF
		sqlStr = sqlStr + " and itemoption=" + itemoptionarr(i)  + VbCrlF

		rsget.Open sqlStr,dbget,1
	next
end if
%>
<script language="javascript">
alert('저장 되었습니다.');
location.replace('brandjaegolist.asp?shopid=<%= shopid %>&makerid=<%= designer %>');
</script>

<!-- #include virtual="/lib/db/dbclose.asp" -->
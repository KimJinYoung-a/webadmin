<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->

<%
dim mode
dim code, ipchulflag
dim itemgubun, itemid, itemoption
dim newitemgubun, newitemid, newitemoption
dim newitemname, newitemoptionname
dim ipchuldetailid, sheetdetailid
dim oldnnew, detailidx, orderserial

dim refer
refer = request.ServerVariables("HTTP_REFERER")

mode = request("mode")
code = request("code")
ipchulflag = request("ipchulflag")
itemgubun = request("itemgubun")
itemid = request("itemid")
itemoption = request("itemoption")
newitemgubun = request("newitemgubun")
newitemid = request("newitemid")
newitemoption   = request("newitemoption")
newitemname = html2db(request("newitemname"))
newitemoptionname = html2db(request("newitemoptionname"))

ipchuldetailid	= request("ipchuldetailid")
sheetdetailid = request("sheetdetailid")

oldnnew = request("oldnnew")
detailidx = request("detailidx")
orderserial = request("orderserial")
dim sqlstr,i

if mode="editipchuldetailwithjungsan" then
	''1 입출고 내역 수정
	sqlstr = "update [db_storage].[dbo].tbl_acount_storage_detail"
	sqlstr = sqlstr + " set iitemgubun='" + newitemgubun + "'"
	sqlstr = sqlstr + " ,itemid=" + CStr(newitemid)
	sqlstr = sqlstr + " ,itemoption='" + newitemoption + "'"
	sqlstr = sqlstr + " ,iitemname='" + newitemname + "'"
	sqlstr = sqlstr + " ,iitemoptionname='" + newitemoptionname + "'"
	sqlstr = sqlstr + " ,updt=getdate()"
	sqlstr = sqlstr + " where mastercode='" + code + "'"
	sqlstr = sqlstr + " and id=" + CStr(ipchuldetailid)

	rsget.Open sqlStr,dbget,1


	''정산내역
	if ipchulflag="I" then
		''입고인경우 - 매입정산부분 수정
		sqlstr = " update [db_jungsan].[dbo].tbl_designer_jungsan_detail"
		sqlstr = sqlstr + " set itemid=" + CStr(newitemid)
		sqlstr = sqlstr + " , itemoption='" + CStr(newitemoption) + "'"
		sqlstr = sqlstr + " , itemname='" + CStr(newitemname) + "'"
		sqlstr = sqlstr + " , itemoptionname='" + CStr(newitemoptionname) + "'"
		sqlstr = sqlstr + " where mastercode='" + code + "'"
		sqlstr = sqlstr + " and detailidx=" + CStr(ipchuldetailid)

		rsget.Open sqlStr,dbget,1

	elseif ipchulflag="S" then
		''샾출고인경우
		sqlstr = "update [db_shop].[dbo].tbl_shop_jungsandetail"
		sqlstr = sqlstr + " set itemgubun='" + newitemgubun + "'"
		sqlstr = sqlstr + " ,itemid=" + CStr(newitemid)
		sqlstr = sqlstr + " ,itemoption='" + newitemoption + "'"
		sqlstr = sqlstr + " ,itemname='" + newitemname + "'"
		sqlstr = sqlstr + " ,itemoptionname='" + newitemoptionname + "'"
		sqlstr = sqlstr + " where orderno='" + code + "'"
		sqlstr = sqlstr + " and linkidx=" + CStr(ipchuldetailid)

		rsget.Open sqlStr,dbget,1
	elseif ipchulflag="E" then
		''기타 출고인경우
		sqlstr = " update [db_jungsan].[dbo].tbl_designer_jungsan_detail"
		sqlstr = sqlstr + " set itemid=" + CStr(newitemid)
		sqlstr = sqlstr + " , itemoption='" + CStr(newitemoption) + "'"
		sqlstr = sqlstr + " , itemname='" + CStr(newitemname) + "'"
		sqlstr = sqlstr + " , itemoptionname='" + CStr(newitemoptionname) + "'"
		sqlstr = sqlstr + " where mastercode='" + code + "'"
		sqlstr = sqlstr + " and detailidx=" + CStr(ipchuldetailid)

		rsget.Open sqlStr,dbget,1
	end if

elseif mode="editordersheetdetail" then
	sqlstr = "update [db_storage].[dbo].tbl_ordersheet_detail"
	sqlstr = sqlstr + " set itemgubun='" + newitemgubun + "'"
	sqlstr = sqlstr + " ,itemid=" + CStr(newitemid)
	sqlstr = sqlstr + " ,itemoption='" + newitemoption + "'"
	sqlstr = sqlstr + " ,itemname='" + newitemname + "'"
	sqlstr = sqlstr + " ,itemoptionname='" + newitemoptionname + "'"
	sqlstr = sqlstr + " ,updt=getdate()"
	sqlstr = sqlstr + " where idx=" + CStr(sheetdetailid)

	rsget.Open sqlStr,dbget,1

elseif mode="editipchulsheetmeaipjungsan" then
	''1 입출고 내역 수정
	sqlstr = "update [db_storage].[dbo].tbl_acount_storage_detail"
	sqlstr = sqlstr + " set iitemgubun='" + newitemgubun + "'"
	sqlstr = sqlstr + " ,itemid=" + CStr(newitemid)
	sqlstr = sqlstr + " ,itemoption='" + newitemoption + "'"
	sqlstr = sqlstr + " ,iitemname='" + newitemname + "'"
	sqlstr = sqlstr + " ,iitemoptionname='" + newitemoptionname + "'"
	sqlstr = sqlstr + " ,updt=getdate()"
	sqlstr = sqlstr + " where mastercode='" + code + "'"
	sqlstr = sqlstr + " and id=" + CStr(ipchuldetailid)

	rsget.Open sqlStr,dbget,1

	''주문서 내역
	if (ipchulflag="S") and (sheetdetailid<>"") then
		sqlstr = "update [db_storage].[dbo].tbl_ordersheet_detail"
		sqlstr = sqlstr + " set itemgubun='" + newitemgubun + "'"
		sqlstr = sqlstr + " ,itemid=" + CStr(newitemid)
		sqlstr = sqlstr + " ,itemoption='" + newitemoption + "'"
		sqlstr = sqlstr + " ,itemname='" + newitemname + "'"
		sqlstr = sqlstr + " ,itemoptionname='" + newitemoptionname + "'"
		sqlstr = sqlstr + " where idx=" + CStr(ordersheetdetailid)

		rsget.Open sqlStr,dbget,1
	end if


	''정산내역
	if ipchulflag="I" then
		''입고인경우 - 매입정산부분 수정
		sqlstr = " update from [db_jungsan].[dbo].tbl_designer_jungsan_detail"
		sqlstr = sqlstr + " set itemid=" + CStr(newitemid)
		sqlstr = sqlstr + " , itemoption='" + CStr(newitemoption) + "'"
		sqlstr = sqlstr + " , itemname='" + CStr(newitemname) + "'"
		sqlstr = sqlstr + " , itemoptionname='" + CStr(newitemoptionname) + "'"
		sqlstr = sqlstr + " where mastercode='" + code + "'"
		sqlstr = sqlstr + " and detailidx=" + CStr(ipchuldetailid)

		rsget.Open sqlStr,dbget,1

	elseif ipchulflag="S" then
		''샾출고인경우
		sqlstr = "update [db_shop].[dbo].tbl_shop_jungsandetail"
		sqlstr = sqlstr + " set itemgubun='" + newitemgubun + "'"
		sqlstr = sqlstr + " ,itemid=" + CStr(newitemid)
		sqlstr = sqlstr + " ,itemoption='" + newitemoption + "'"
		sqlstr = sqlstr + " ,itemname='" + newitemname + "'"
		sqlstr = sqlstr + " ,itemoptionname='" + newitemoptionname + "'"
		sqlstr = sqlstr + " where orderno='" + code + "'"
		sqlstr = sqlstr + " and linkidx=" + CStr(ipchuldetailid)

		rsget.Open sqlStr,dbget,1
	elseif ipchulflag="E" then
		''기타 출고인경우
		sqlstr = " update from [db_jungsan].[dbo].tbl_designer_jungsan_detail"
		sqlstr = sqlstr + " set itemid=" + CStr(newitemid)
		sqlstr = sqlstr + " , itemoption='" + CStr(newitemoption) + "'"
		sqlstr = sqlstr + " , itemname='" + CStr(newitemname) + "'"
		sqlstr = sqlstr + " , itemoptionname='" + CStr(newitemoptionname) + "'"
		sqlstr = sqlstr + " where mastercode='" + code + "'"
		sqlstr = sqlstr + " and detailidx=" + CStr(ipchuldetailid)

		rsget.Open sqlStr,dbget,1
	end if

elseif mode="editselldetail" then

	if oldnnew="old" then
		sqlstr = "update [db_log].[dbo].tbl_old_order_detail_2003"
		sqlstr = sqlstr + " set itemid=" + CStr(newitemid)
		sqlstr = sqlstr + " ,itemoption='" + CStr(newitemoption) + "'"
		sqlstr = sqlstr + " ,itemname='" + CStr(newitemname) + "'"
		sqlstr = sqlstr + " ,itemoptionname='" + CStr(newitemoptionname) + "'"
		sqlstr = sqlstr + " where idx=" + CStr(detailidx)

		rsget.Open sqlStr,dbget,1
	else
		sqlstr = "update [db_order].[dbo].tbl_order_detail"
		sqlstr = sqlstr + " set itemid=" + CStr(newitemid)
		sqlstr = sqlstr + " ,itemoption='" + CStr(newitemoption) + "'"
		sqlstr = sqlstr + " ,itemname='" + CStr(newitemname) + "'"
		sqlstr = sqlstr + " ,itemoptionname='" + CStr(newitemoptionname) + "'"
		sqlstr = sqlstr + " where idx=" + CStr(detailidx)

		rsget.Open sqlStr,dbget,1
	end if



	sqlstr = "update [db_jungsan].[dbo].tbl_designer_jungsan_detail"
	sqlstr = sqlstr + " set itemid=" + CStr(newitemid)
	sqlstr = sqlstr + " ,itemoption='" + CStr(newitemoption) + "'"
	sqlstr = sqlstr + " ,itemname='" + CStr(newitemname) + "'"
	sqlstr = sqlstr + " ,itemoptionname='" + CStr(newitemoptionname) + "'"
	sqlstr = sqlstr + " where detailidx=" + CStr(detailidx)
	sqlstr = sqlstr + " and mastercode='" + orderserial + "'"

	rsget.Open sqlStr,dbget,1
end if
%>
<script language="javascript">
alert('저장 되었습니다.');
location.replace('<%= refer %>');
</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->
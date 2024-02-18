<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->

<%
'response.write " 수정중 "
'dbget.close()	:	response.End

dim mode
dim itemidarr
dim refer
refer = request.ServerVariables("HTTP_REFERER")

mode = RequestCheckvar(request.form("mode"),16)
itemidarr = trim(request.form("itemidarr"))
if itemidarr <> "" then
	if checkNotValidHTML(itemidarr) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');history.back();"
	response.write "</script>"
	response.End
	end if
end if
if Right(itemidarr,1)="," then itemidarr=Left(itemidarr,Len(itemidarr)-1)


dim sqlStr, addRows

if mode="addArr" then
	''상품동기화
	
	sqlStr = "insert into [db_item].[dbo].tbl_academy_product" + VbCrlf
	sqlStr = sqlStr + "(itemid, reguserid)" + VbCrlf
	sqlStr = sqlStr + " select i.itemid,'" + session("ssBctId") + "'" + VbCrlf
	sqlStr = sqlStr + " from [db_item].[dbo].tbl_item i" + VbCrlf
	sqlStr = sqlStr + " left join [db_item].[dbo].tbl_academy_product p on i.itemid=p.itemid" + VbCrlf
	sqlStr = sqlStr + " where i.itemid in (" + itemidarr + ")" + VbCrlf
	sqlStr = sqlStr + " and p.itemid is null"

	dbget.Execute  sqlStr, addRows
	
	
	sqlStr = " insert into [db_academy].[dbo].tbl_academy_product " + VbCrlf
	sqlStr = sqlStr + "(itemid, regdate, reguserid)" + VbCrlf
	sqlStr = sqlStr + " select p.itemid, p.regdate, p.reguserid  "
	sqlStr = sqlStr + " from [110.93.128.72].[db_item].[dbo].tbl_academy_product p"
	sqlStr = sqlStr + " left join [db_academy].[dbo].tbl_academy_product t on p.itemid=t.itemid"
	sqlStr = sqlStr + " where p.itemid in (" + itemidarr + ")" + VbCrlf
	sqlStr = sqlStr + " and t.itemid is null"
	
	dbAcademyget.Execute  sqlStr, addRows
	
	'response.write addRows
	'dbget.close()	:	response.End
	
elseif mode="dellarr" then
	if itemidarr="" then
		response.write "<script>alert('선택되 상품이 없습니다.');</script>"
		response.write "<script>location.replace('" + refer + "');</script>"
		dbget.close()	:	response.End
	end if

	sqlStr = "delete from [db_item].[dbo].tbl_academy_product" + VbCrlf
	sqlStr = sqlStr + " where itemid in (" + itemidarr + ")" + VbCrlf

	dbget.Execute  sqlStr, addRows
	
	sqlStr = "delete from [db_academy].[dbo].tbl_academy_product" + VbCrlf
	sqlStr = sqlStr + " where itemid in (" + itemidarr + ")" + VbCrlf

	dbAcademyget.Execute  sqlStr, addRows
elseif mode ="bestarr" then	
	if itemidarr="" then
		response.write "<script>alert('선택되 상품이 없습니다.');</script>"
		response.write "<script>location.replace('" + refer + "');</script>"
		dbget.close()	:	response.End
	end if	
	
	sqlStr = "update [db_item].[dbo].tbl_academy_product set isBest ='Y' " + VbCrlf
	sqlStr = sqlStr + " where itemid in (" + itemidarr + ")" + VbCrlf

	dbget.Execute  sqlStr, addRows
	
	sqlStr = "update [db_academy].[dbo].tbl_academy_product set isBest ='Y' " + VbCrlf
	sqlStr = sqlStr + " where itemid in (" + itemidarr + ")" + VbCrlf

	dbAcademyget.Execute  sqlStr, addRows

elseif mode ="unbest"	then
	Dim bestId
	bestId = request("bestId")
	
	sqlStr = "update [db_item].[dbo].tbl_academy_product set isBest ='N' " + VbCrlf
	sqlStr = sqlStr + " where itemid = " + bestId +  VbCrlf

	dbget.Execute  sqlStr, addRows
	
	sqlStr = "update [db_academy].[dbo].tbl_academy_product set isBest ='N' " + VbCrlf
	sqlStr = sqlStr + " where itemid = " + bestId + VbCrlf

	dbAcademyget.Execute  sqlStr, addRows
else

end if

%>
<script language="javascript">
alert('<%= CStr(addRows) %> 행 적용 되었습니다.');
location.replace('<%= refer %>');
</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->
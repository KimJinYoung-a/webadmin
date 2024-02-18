<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<%
dim mode,itemid,sortNo,cdl, cdm, vIdx
dim viewidx,disptitle,allusing
dim i, itemsYN

mode = request("mode")
cdl = request("cdl")
cdm = request("cdm")
itemid = Trim(Replace(request("itemid")," ",""))
sortNo = request("sortNo")
viewidx = request("viewidx")
disptitle = request("disptitle")
allusing = request("allusing")
vIdx = request("idx")
itemsYN = request("itemsYN")

'// 전송된 아이템 코드값 확인
if Right(itemid,1)="," then
	itemid = Left(itemid,Len(itemid)-1)
end if

dim sqlStr,msg

on error resume  next 

dbget.BeginTrans

if mode="del" then
	sqlStr = "delete from [db_sitemaster].[dbo].tbl_category_MDChoice where itemid in (" + itemid + ") and theme_idx = '" & vIdx & "' "

elseif mode="add" then
	If itemsYN = "Y" Then
		dim iA ,arrTemp,arrItemid
		itemid = request("arrItems")
		itemid = replace(itemid,",",chr(10))
		itemid = replace(itemid,chr(13),"")
		arrTemp = Split(itemid,chr(10))
	
		iA = 0
		do while iA <= ubound(arrTemp) 
			if trim(arrTemp(iA))<>"" then
				'상품코드 유효성 검사(2008.08.05;허진원)
				if Not(isNumeric(trim(arrTemp(iA)))) then
					Response.Write "<script language=javascript>alert('[" & arrTemp(iA) & "]은(는) 유효한 상품코드가 아닙니다.');history.back();</script>"
					dbget.close()	:	response.End
				else
					arrItemid = arrItemid & trim(arrTemp(iA)) & ","
				end if
			end if
			iA = iA + 1
		loop
	
		itemid = left(arrItemid,len(arrItemid)-1)
	End If

	sqlStr = "insert into [db_sitemaster].[dbo].tbl_category_MDChoice" &_
				" (theme_idx, itemid, sortNo)" &_
				" select '" & vIdx & "', itemid, '99' " &_
				" from [db_item].[dbo].tbl_item" &_
				" where itemid in (" & itemid & ")" 
elseif mode="isUsingValue" then
	sqlStr = "update [db_sitemaster].[dbo].tbl_category_MDChoice set isusing = '" & allusing & "' where itemid in (" & itemid & ") and theme_idx = '" & vIdx & "'"

elseif mode="ChangeSort" then
	itemid = split(itemid,",")
	sortNo = split(sortNo,",")
	sqlStr = ""
	for i=0 to ubound(itemid)
		sqlStr = sqlStr & " update [db_sitemaster].[dbo].tbl_category_MDChoice " &_
					" set sortNo='" & sortNo(i) & "'" &_
					" where itemid='" & itemid(i) & "' and theme_idx = '" & vIdx & "';" & vbCrLf
	next

end if

dbget.execute(sqlStr)


if err.number<>0 then
	dbget.rollback
	msg ="오류 발생, 관리자문의 요망"
else
	dbget.committrans
	msg ="적용 되었습니다."
end if
dim refer
refer = request.ServerVariables("HTTP_REFERER")
%>
<script language="javascript">
//alert('<%= msg %>');
//opener.location.reload();
location.replace('<%= refer %>');
</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->

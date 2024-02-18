<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->

<%
dim diaryid,mode,itemid,sql
dim arritemid,arrcnt,i
dim referer
referer = request.ServerVariables("HTTP_REFERER")

diaryid = request("diaryid")
mode=request("mode")
itemid=request("itemid")


if mode="write" then
	if not (itemid="" or isnull(itemid)) then

		arritemid=split(itemid,",")
		arrcnt=ubound(arritemid)

		for i=0 to arrcnt
		sql = sql + " INSERT INTO [db_diary_collection].[dbo].tbl_diary_linkitem (idx,itemid) " &_
								" VALUES(" & diaryid & "," & arritemid(i) & ")"
		next

		rsget.open sql,dbget,1

	end if

elseif mode="del" then

	if not (itemid="" or isnull(itemid)) then
		sql	=	" DELETE FROM [db_diary_collection].[dbo].tbl_diary_linkitem " &_
					" WHERE idx='" & diaryid & "'" &_
					" and itemid in (" & itemid &")"

	rsget.open sql,dbget,1
	end if
end if

response.write "<script language='javascript'>alert('적용되었습니다.')</script>"
response.write "<script language='javascript'>location.replace('" + CStr(referer) + "')</script>"
dbget.close()	:	response.End


%>

<!-- #include virtual="/lib/db/dbclose.asp" -->
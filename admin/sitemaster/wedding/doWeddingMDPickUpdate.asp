<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<%
Dim idx, DispOrder, mode
Dim arrSort, arrIdx, lp, itemarr
	
	idx = requestCheckVar(Replace(Request("selIdx")," ",""),128)
	DispOrder = requestCheckVar(Replace(Request("DispOrder")," ",""),128)
	itemarr = requestCheckVar(Replace(Request("itemarr")," ",""),128)
	arrSort = requestCheckVar(Replace(Request("arrSort")," ",""),128)
	mode = requestCheckVar(request("mode"),16)
dim sqlStr

if mode = "changeUsing" then
   sqlStr = " update [db_sitemaster].[dbo].[tbl_wedding_md_pick]" + VbCrlf
   sqlStr = sqlStr + " set isusing='N'" + VbCrlf
   sqlStr = sqlStr + " where idx in (" + cstr(idx) + ")"
   dbget.Execute sqlStr
Elseif mode = "changeSort" Then
	If idx<>"" Then
		arrIdx = split(idx,",")
		arrSort = split(arrSort,",")
		For lp=0 To ubound(arrIdx)
		   sqlStr = " update [db_sitemaster].[dbo].[tbl_wedding_md_pick]" + VbCrlf
		   sqlStr = sqlStr + " set DispOrder='" + Cstr(arrSort(lp)) + "'" + VbCrlf
		   sqlStr = sqlStr + " where idx ='" + cstr(arrIdx(lp)) + "'"
		   dbget.Execute sqlStr
		Next
	End If
Elseif mode = "multi" Then
	If itemarr<>"" Then
		arrIdx = split(itemarr,",")

		For lp=0 To ubound(arrIdx)
		   sqlStr = " insert into [db_sitemaster].[dbo].[tbl_wedding_md_pick](itemid, LastUser, DispOrder)" + VbCrlf
		   sqlStr = sqlStr + " values('" + Cstr(arrIdx(lp)) + "'" + VbCrlf
		   sqlStr = sqlStr + " , '" + cstr(session("ssBctCname")) + "'"
		   sqlStr = sqlStr + " , '" + cstr(lp) + "')"
		   dbget.Execute sqlStr
		Next
	End If
End If

dim referer
	referer = request.ServerVariables("HTTP_REFERER")
	response.write "<script>alert('저장되었습니다.');location.href='/admin/sitemaster/wedding/md_pick_manager.asp'</script>"
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<%
dim mode, idx
dim idxarr,viewidxarr

mode = request.Form("mode")
idxarr = request.Form("idxarr")
viewidxarr = request.Form("viewidxarr")

dim sqlStr, i, cnt

if (mode="viewidxedit") then

	idxarr = Left(idxarr,Len(idxarr)-1)
	viewidxarr = Left(viewidxarr,Len(viewidxarr)-1)

	idxarr = split(idxarr,"|")
	viewidxarr = split(viewidxarr,"|")

	cnt = ubound(idxarr)

	for i=0 to cnt
		sqlStr = " update [db_contents].[dbo].tbl_lecture_item" + VbCrlf
		sqlStr = sqlStr + " set viewidx=" + viewidxarr(i) + "" + VbCrlf
		sqlStr = sqlStr + " where  idx=" + idxarr(i) + VbCrlf
		'response.write sqlStr
		rsget.Open sqlStr,dbget,1
	next

end if

dim refer
refer = request.ServerVariables("HTTP_REFERER")

%>

<script language="javascript">
alert('적용 되었습니다.');
location.replace('<%= refer %>');
</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->
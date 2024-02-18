<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  이미지 관리
' History : 2016.08.12 한용민 수정
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/items/overseas/overseasCls.asp"-->
<%
dim sqlStr, userid, menupos, mode, folderidx, folderTitle, realPath, sortkey, isusing
	folderidx   = getNumeric(requestcheckvar(request("folderidx"),10))
	folderTitle   = requestcheckvar(request("folderTitle"),60)
	realPath   = requestcheckvar(request("realPath"),200)
	sortkey   = getNumeric(requestcheckvar(request("sortkey"),10))
	isusing   = requestcheckvar(request("isusing"),1)
	mode   = requestcheckvar(request("mode"),32)

if (sortkey="") then sortkey="100"

dim referer
	referer = request.ServerVariables("HTTP_REFERER")

if mode = "" then
	response.write "<script type='text/javascript'>"
	response.write "	alert(MODE 구분자가 지정되지 않았습니다.');"
	response.write "	location.replace('" & referer & "');"
	response.write "</script>"
end if

if mode = "etcimgedit" then
	sqlStr = "if exists(" + VbCrlf
	sqlStr = sqlStr & "		select top 1 * from db_event.[dbo].[tbl_etcImage_master] where folderIdx='"& folderIdx &"'" + VbCrlf
	sqlStr = sqlStr & " )" + VbCrlf
    sqlStr = sqlStr & " 	update db_event.[dbo].[tbl_etcImage_master]" + VbCrlf
    sqlStr = sqlStr & " 	set folderTitle=N'" + html2db(folderTitle) + "'" + VbCrlf
    sqlStr = sqlStr & " 	,realPath=N'" + realPath + "'" + VbCrlf
    sqlStr = sqlStr & " 	,sortkey=" + sortkey + "" + VbCrlf
    sqlStr = sqlStr & " 	,isusing=N'" + isusing + "'" + VbCrlf
    sqlStr = sqlStr & " 	where folderIdx=N'" + folderIdx + "'" + VbCrlf
	sqlStr = sqlStr & " else" + VbCrlf
	sqlStr = sqlStr & " 	insert into db_event.[dbo].[tbl_etcImage_master] (" + VbCrlf
    sqlStr = sqlStr & " 	folderTitle, realPath, sortkey, isusing"+ VbCrlf
	sqlStr = sqlStr & " 	) values("
    sqlStr = sqlStr & " 	N'" + html2db(folderTitle) + "'" + VbCrlf
    sqlStr = sqlStr & " 	,N'" + realPath + "'" + VbCrlf
    sqlStr = sqlStr & " 	," + sortkey + "" + VbCrlf
    sqlStr = sqlStr & " 	,N'" + isusing + "'" + VbCrlf
    sqlStr = sqlStr & " 	)"

	'response.write sqlStr &"<Br>"
    dbget.Execute sqlStr

elseif mode = "etcimgdel" then
	sqlStr = "delete from db_event.[dbo].[tbl_etcImage_master]" + VbCrlf
	sqlStr = sqlStr & " where folderIdx='"& folderIdx &"'"

	'response.write sqlStr &"<Br>"
    dbget.Execute sqlStr
end if
%>

<script language='javascript'>
	alert('OK');
	location.replace('<%=referer%>');
</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<% session.codePage = 949 %>
<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/classes/offshop_noticecls.asp" -->
<%

'// 사용안함, 2015-07-07, skyer9
dbget.close()
response.end

Function getFileExt(str)
	dim sp
	sp = split(str,".")
	getFileExt = sp(UBound(sp))
End Function

dim boardnotice
dim boarditem
dim idx, mode, title, contents, userid, username, gubun, file, dl_file

dim uploadForm,objFSO
Set uploadForm = Server.CreateObject("DEXT.FileUpload")
Set objFSO = Server.CreateObject("Scripting.FileSystemObject")

idx = uploadForm("idx")
gubun = uploadForm("gubun")
mode = uploadForm("mode")
title = html2db(uploadForm("title"))
contents = html2db(uploadForm("contents"))
userid = uploadForm("userid")
username = uploadForm("username")
file = uploadForm("file")
dl_file = uploadForm("dl_file")

dim updir
updir = Server.MapPath("\admin\board\noticefile\")

dim sql,iid

if (mode = "write") then

	sql = " insert into [db_board].[10x10].tbl_offshop_notice(gubun,userid,username,title,contents) "
	sql = sql + " values('" + gubun + "', '" + userid + "', '" + username + "', '" + title + "','" + contents + "') "
	rsget.Open sql, dbget, 1

	sql = " select IDENT_CURRENT('[db_contents].[dbo].[tbl_finger_master]') as currid"
	rsget.Open sql,dbget,1
	iid = rsget("currid")
	rsget.close

elseif (mode = "modify") then

	sql = "update [db_board].[10x10].tbl_offshop_notice " + VbCrlf
	sql = sql + " set gubun = '" + gubun + "'," + VbCrlf
	sql = sql + " userid = '" + userid + "'," + VbCrlf
	sql = sql + " username = '" + username + "'," + VbCrlf
	sql = sql + " title = '" + title + "'," + VbCrlf
	sql = sql + " contents = '" + contents + "'" + VbCrlf
	sql = sql + " where (idx = " + idx + ") "
	rsget.Open sql, dbget, 1

	iid = idx
elseif (mode = "delete") then

	sql = "update [db_board].[10x10].tbl_offshop_notice set isusing = 'N' " + VbCrlf
	sql = sql + " where (idx = " + idx + ") "
	rsget.Open sql, dbget, 1

	'response.replace("offshop_notice_list.asp")
end if

function DelProc(fieldname, idx)

	sql = " update [db_board].[10x10].tbl_offshop_notice"
	sql = sql + " set " + fieldname + "=NULL"
	sql = sql + " where idx=" + CStr(idx)
	rsget.Open sql,dbget,1

end function


dim filenameolny, svrname_img


if (file<>"") and (dl_file<>"on") then

	filenameolny =  "notice_file" + Format00(7,iid) + "." + getFileExt(file)
	svrname_img = updir & "\" & filenameolny

	uploadForm("file").saveas(svrname_img)


	sql = " update [db_board].[10x10].tbl_offshop_notice"
	sql = sql + " set  [file] ='" + filenameolny + "'"
	sql = sql + " where idx=" + CStr(iid)
	rsget.Open sql,dbget,1
end if

if (dl_file="on") then
	DelProc "[file]", iid
end if

Set uploadForm = Nothing
Set objFSO = Nothing

dim refer
refer = request.ServerVariables("HTTP_REFERER")
%>

<script language="javascript">
alert('저장 되었습니다.');
location.replace('<%= refer %>');
</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->

<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  ��ġ����Ŀ ���� ������ Iframe�̹������ ó�� ������
' History : 2014.08.04 ���¿� ����
'###########################################################
%>
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->

<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/classes/hitchhiker/hitchhiker_previewCls.asp"-->
<%
Dim tmpSort, tmpIsusing
Dim cnt, i, sqlStr, idx, mode
Dim detailidxarr, isusingarr, sortnoarr, device
	idx	= Request("idx")
	mode = Request("mode")
	sortnoarr 	= Request("sortnoarr")
	isusingarr = Request("isusingarr")
	detailidxarr = Request("detailidxarr")
	device		= Request("device")

if mode="sortisusingedit" then

	'�����̹��� �ľ�
	detailidxarr = split(detailidxarr,",")
	cnt = ubound(detailidxarr)
	
	sortnoarr	=  split(sortnoarr,",")
	isusingarr	=  split(isusingarr,",")
	
	For i = 0 to cnt
		tmpSort = sortnoarr(i)
		tmpIsusing = isusingarr(i)
		
		sqlStr = ""
		sqlStr = sqlStr & " UPDATE db_sitemaster.dbo.tbl_hitchhiker_preview_Detail SET " & VBCRLF
		sqlStr = sqlStr & " isusing = '"&tmpIsusing&"'" & VBCRLF
		sqlStr = sqlStr & " ,sortnum = '"&tmpSort&"'" & VBCRLF
		sqlStr = sqlStr & " WHERE detailidx =" & detailidxarr(i)
		
		'response.write sqlStr & "<Br>"
		dbget.execute sqlStr
	Next

	response.write "<script language='javascript'>"
	response.write "	alert('����Ǿ����ϴ�');"
	if device = "W" then
		response.write "	location.replace('/admin/hitchhiker/preview/iframe_hitchhiker_preview.asp?idx="&idx&"');"
	else
		response.write "	location.replace('/admin/hitchhiker/preview/iframe_hitchhiker_preview_M.asp?idx="&idx&"');"
	end if
	response.write "</script>"
else
	Response.Write "<script language='javascript'>alert('�����ڰ� �����ϴ�.'); history.back(-1);</script>"
	dbget.close()	:	response.End	
end if
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  �귣�彺Ʈ��Ʈ
' History : 2013.08.29 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/street/interviewCls.asp"-->
<%
Dim detailidxarr, sortnoarr, tmpSort, tmpIsusing, cnt, i, sqlStr, isusingarr, idx, mode, adminid
	sortnoarr 	= Request("sortnoarr")
	detailidxarr 	= Request("detailidxarr")
	isusingarr	= Request("isusingarr")
	idx			= Request("idx")
	mode 		= Request("mode")

adminid = session("ssBctId")

if mode="sortisusingedit" then
	If sortnoarr="" THEN
		Response.Write "<script language='javascript'>history.back(-1);</script>"
		dbget.close()	:	response.End
	end if
	
	'���û�ǰ �ľ�
	detailidxarr = split(detailidxarr,",")
	cnt = ubound(detailidxarr)
	
	sortnoarr 	=  split(sortnoarr,",")
	isusingarr	=  split(isusingarr,",")
	
	For i = 0 to cnt
		tmpSort = sortnoarr(i)	
		tmpIsusing = isusingarr(i)
	
		sqlStr = ""
		sqlStr = sqlStr & " UPDATE db_brand.dbo.tbl_street_interview_detail SET " & VBCRLF
		sqlStr = sqlStr & " sortNo = '"&tmpSort&"'" & VBCRLF
		sqlStr = sqlStr & " ,isusing = '"&tmpIsusing&"'" & VBCRLF
		sqlStr = sqlStr & " ,lastupdate=getdate()" + vbcrlf
		sqlStr = sqlStr & " ,lastadminid = '"&adminid&"'" + vbcrlf		
		sqlStr = sqlStr & " WHERE detailidx =" & detailidxarr(i)
		
		'response.write sqlStr & "<Br>"
		dbget.execute sqlStr
	Next

	response.write "<script language='javascript'>"
	response.write "	alert('����Ǿ����ϴ�');"
	response.write "	location.replace('/admin/brand/interview/iframe_interview_detail.asp?idx="&idx&"');"
	response.write "</script>"

else
	Response.Write "<script language='javascript'>alert('�����ڰ� �����ϴ�.'); history.back(-1);</script>"
	dbget.close()	:	response.End	
end if
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
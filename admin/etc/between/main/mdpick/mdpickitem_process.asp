<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbCTopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<%
Dim mode, idxarr, page, gender
Dim sqlStr
mode		= Request("mode")
idxarr		= Request("idxarr")
page		= request("page")
gender		= request("gender")

dbCTget.beginTrans
Select Case mode
	Case "S" '//��ǰ���� ����
		Dim tmpSort, sortarr, cnt, i
		sortarr = Request("sortarr")

		If sortarr="" Then
			dbCTget.RollBackTrans
			Response.Write "<script language='javascript'>history.back(-1);</script>"
			dbCTget.close()	:	response.End
		End if

		'���û�ǰ �ľ�
		idxarr = split(idxarr,",")
		cnt = ubound(idxarr)

		'// ���ļ��� ����
		If sortarr<>"" THEN
			sortarr =  split(sortarr,",")

			For i = 0 to cnt
				If sortarr(i) = "" Then
					 tmpSort = "NULL"
				Else
					 tmpSort = sortarr(i)
				End If
				sqlStr = ""
				sqlStr = sqlStr & " UPDATE db_outmall.dbo.tbl_between_project_groupItem SET "
				sqlStr = sqlStr & " MainMdpickSortNo = '"&tmpSort&"'"
				sqlStr = sqlStr & "	WHERE idx =" & idxarr(i)
				dbCTget.execute sqlStr
			Next
		End If
End Select

IF Err.Number = 0 THEN
	dbCTget.CommitTrans
		response.redirect("index.asp?menupos="&menupos&"&page="&page&"&gender="&gender)
	dbCTget.close()	:	response.End
Else
   	dbCTget.RollBackTrans
%>
	<script language="javascript">
	alert("������ ó���� ������ �߻��Ͽ����ϴ�.");
	history.back(-1);
	</script>
<%
	dbCTget.close()	:	response.End
End IF
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbCTclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->
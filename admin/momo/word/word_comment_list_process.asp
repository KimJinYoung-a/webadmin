<%@ language=vbscript %>
<% option explicit %>

<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/momo/momo_cls.asp"-->

<%
	Dim vQuery, vTotIdx, vIsBest, vNullIdx, vPage, vKeyid, i, vGubun
	vTotIdx	= Request("totidx")
	vTotIdx = Left(vTotIdx,(Len(vTotIdx)-1))
	vNullIdx = vTotIdx
	vIsBest = Request("isbest")
	vPage	= Request("nowpage")
	vKeyid	= Request("keyid")
	vGubun	= Request("gubun")
	
	For i=0 To UBound(Split(vIsBest,","))
		vNullIdx = Replace(vNullIdx,Trim(Split(vIsBest,",")(i)),"")
		vNullIdx = Replace(vNullIdx,",,",",")
	Next

	IF Left(vNullIdx,1) = "," Then
		vNullIdx = Right(vNullIdx,Len(vNullIdx)-1)
	End If
	
	IF Right(vNullIdx,1) = "," Then
		vNullIdx = Left(vNullIdx,Len(vNullIdx)-1)
	End If

	On Error Resume Next
	dbget.beginTrans
	
	
	'### üũ �����Ȱ� Null �� ������Ʈ
	IF vNullIdx <> "" Then
		vQuery = "UPDATE [db_momo].[dbo].[tbl_word_comment] SET isbest = Null WHERE idx IN(" & vNullIdx & ") "
		dbget.execute vQuery
	End IF
	
	'response.write vQuery & "<br>"
	'### üũ �Ȱ� isbest = 'o' �� ������Ʈ
	IF vIsBest <> "" Then
		vQuery = "UPDATE [db_momo].[dbo].[tbl_word_comment] SET isbest = 'o' WHERE idx IN(" & vIsBest & ") "
		dbget.execute vQuery
	End IF
	
	'response.write vQuery
	If Err.Number = 0 Then
	        dbget.CommitTrans
	Else
	        dbget.RollBackTrans
	        dbget.close()
	        response.write "<script>alert('����Ÿ�� �����ϴ� ���߿� ������ �߻��Ͽ����ϴ�.');history.back();</script>"
	        response.end
	End If
	on error Goto 0
	
    response.write "<script language='javascript'>alert('����Ǿ����ϴ�.');location.href='/admin/momo/word/word_comment_list.asp?keyid="&vKeyid&"&page="&vPage&"&gubun="&vGubun&"';</script>"
    dbget.close()
    response.end
%>

<!-- #include virtual="/common/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
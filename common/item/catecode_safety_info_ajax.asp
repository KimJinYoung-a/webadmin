<%@ language=vbscript %>
<% option explicit %>
<% response.charset = "euc-kr" %>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<%
	Dim vCateCode, vQuery, vResultArr, vResult, vResultMsg, i
	vCateCode = Trim(requestCheckVar(request("catecode"),300))
	
	If vCateCode <> "" Then
		If Right(vCateCode,1) = "," Then
			vCateCode = Left(vCateCode, Len(vCateCode)-1)
		End If
		
		vQuery = "select safetyinfotype from db_item.dbo.tbl_display_cate where catecode in(" & vCateCode & ") and safetyinfotype in('necessary','choice')"
		rsget.CursorLocation = adUseClient
		rsget.Open vQuery,dbget,adOpenForwardOnly,adLockReadOnly
		
		If not rsget.eof Then
			Do Until rsget.Eof
				vResultArr = vResultArr & rsget(0) & ","
				rsget.movenext
			Loop
		End If
		rsget.close
		
		If InStr(vResultArr, "necessary") > 0 Then
			vResult = "necessary"
		End If
		
		If vResult <> "necessary" AND InStr(vResultArr, "choice") > 0 Then
				vResult = "choice"
		End If
	End If
	
	If vResult = "necessary" Then
		vResultMsg = "���õ� ī�װ��� �������� ������ �ʼ� �Է� �����Դϴ�.br���� �Է��� �� �� ���br�Ǹ����� �Ǵ� �ǸŻ� �������� ���� �� �ֽ��ϴ�."
	ElseIf vResult = "choice" Then
		vResultMsg = "���õ� ī�װ��� ǰ�� ��br�������� ������ ǰ���� �Ϻ� �����մϴ�.br�������� �ʼ� ǰ���� �ٽ� �ѹ� Ȯ�����ּ���."
	End If
	
	Response.Write vResultMsg
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
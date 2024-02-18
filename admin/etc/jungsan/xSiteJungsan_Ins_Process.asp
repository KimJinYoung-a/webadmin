<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
Server.ScriptTimeOut = 1200 ''�ʴ���
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/JSON_2.0.4.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/util/xmlhttpUtil.asp"-->
<!-- #include virtual="/admin/etc/incOutmallCommonFunction.asp"-->
<!-- #include virtual="/admin/etc/jungsan/lib/xSiteJungsanLib.asp"-->
<script language="jscript" runat="server" src="/lib/util/JSON_PARSER_JS.asp"></script>
<%
Dim sellsite : sellsite	= requestCheckVar(html2db(request("sellsite")),32)
Dim reqDate : reqDate = request("reqDate")
Dim vPage, hasnext, isJungsanComplete, vTotalPage
vPage = request("page")
If vPage = "" Then
	vPage = 1
End If

If sellsite = "ezwel" Then
	If (reqDate = "") Then
		reqDate = Replace(Left(DateAdd("m", -1, NOW()), 7), "-", "")
	End If
	If Len(reqDate) <> 6 Then
		rw "��¥ ������ �� �� �Ǿ����ϴ�."
		response.end
	End If
ElseIf sellsite = "wconcept1010" Then
	reqDate = Left(reqDate,4)&"-"&Mid(reqDate,5,2)&"-"&Mid(reqDate,7,2)
	If Len(reqDate) <> 10 Then
		rw "��¥ ������ �� �� �Ǿ����ϴ�."
		response.end
	End If
Else
	response.write "�߸��� �����Դϴ�."
	dbget.close : response.end
End If

isJungsanComplete = "N"
Select Case sellsite
	Case "ezwel"
		rw "ȣ��� : " & reqDate
		Do Until isJungsanComplete = "Y"
			Call GetJungsan_ezwel(reqDate, hasnext, vPage, vTotalPage)
			If hasnext = "N" Then
				isJungsanComplete = "Y"
				rw "�Ϸ�"
			Else
				rw "API ȣ�� �� �Դϴ�. ("& vPage - 1 & "/" & vTotalPage & ")"
				rw "-------------------------"
			End If
			response.flush
		Loop
	Case "wconcept1010"
		rw "ȣ���� : " & reqDate
		Call GetJungsan_wconcept1010(reqDate, vPage)
		response.flush
		rw "�Ϸ�"
End Select
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->

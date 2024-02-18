<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<%
dim itemid : itemid=123456

Dim vQuery, vIsOK
vQuery = "EXEC [db_item].[dbo].[sp_Ten_ItemNotificationRaw_Check] '" & itemid & "'"
rsget.CursorLocation = adUseClient
'rsget.CursorType = adOpenStatic
'rsget.LockType = adLockOptimistic

rsget.open vQuery,dbget,1
If Not rsget.Eof Then
	vIsOK = rsget(0)
Else
	vIsOK = "x"
End IF
rsget.close()

response.write vIsOK
''----------------------------------------------------------------------------
response.write "<br>"
''----------------------------------------------------------------------------
''------------------------------------------------------------------
vQuery = "[db_item].[dbo].[sp_Ten_ItemNotificationRaw_Check](" & itemid & ")"

rsget.Open vQuery, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
IF Not (rsget.EOF OR rsget.BOF) THEN
	vIsOK = rsget(0)
END IF
rsget.close
		
response.write vIsOK
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
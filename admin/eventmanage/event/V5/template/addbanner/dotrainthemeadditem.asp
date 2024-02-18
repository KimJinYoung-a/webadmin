<%@ language=vbscript %>
<% option explicit %>
<%
Response.Expires = 0   
 Response.AddHeader "Pragma","no-cache"   
 Response.AddHeader "Cache-Control","no-cache,must-revalidate"   

'###########################################################
' Page : dotrainthemeadditem.asp
' Description : I�� �̺�Ʈ ���� ���ø� ������ ���� ���, ����, ����
' History : 2019.02.12 ������
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/eventmanage/common/event_function.asp"-->
<%
'--------------------------------------------------------
' �������� & �Ķ���� �� �ޱ�
'--------------------------------------------------------
Dim k, sqlStr, i
Dim vCnt : vCnt = Request.Form("cksel").count
Dim eC : eC = requestCheckVar(Request.Form("eC"),9)
Dim mode : mode = requestCheckVar(Request.Form("mode"),3)
Dim stype : stype = requestCheckVar(Request.Form("stype"),1)
Dim upback : upback = requestCheckVar(Request.Form("upback"),1)
Dim reUrl : reUrl = Request.ServerVariables("HTTP_REFERER")
Dim GroupItemCheck : GroupItemCheck = requestCheckVar(Request.Form("GroupItemCheck"),1)
dim menuidx : menuidx = requestCheckvar(request("menuidx"),10)

if Request.Form("cksel") <> "" then
	if checkNotValidHTML(Request.Form("cksel")) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('��ȿ���� ���� ���ڰ� ���ԵǾ� �ֽ��ϴ�. �ٽ� �ۼ� ���ּ���');history.back();"
	response.write "</script>"
	response.End
	end if
end if
if Request.Form("sitemname") <> "" then
	if checkNotValidHTML(Request.Form("sitemname")) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('��ȿ���� ���� ���ڰ� ���ԵǾ� �ֽ��ϴ�. �ٽ� �ۼ� ���ּ���');history.back();"
	response.write "</script>"
	response.End
	end if
end If


'�迭�� ó��
redim arritemcode(vCnt)
redim arritemname(vCnt)
for i=1 to vCnt
	arritemcode(i) = Request.Form("cksel")(i)
	arritemname(i) = Request.Form("sitemname")(i)
next
if mode="MR" then
	sqlStr = " Update [db_event].[dbo].[tbl_event_multi_contents_master]"
	sqlStr = sqlStr & " Set GroupItemType='T'"
	sqlStr = sqlStr & " ,GroupItemCheck='" & GroupItemCheck & "'"
	sqlStr = sqlStr & " Where idx=" & menuidx
	dbget.Execute sqlStr
	Response.write "<script>window.document.domain='10x10.co.kr';parent.MainWindowReloadClose();</script>"
	response.End 
ElseIf mode="del" Then
	dbget.beginTrans
			sqlStr = " delete FROM [db_event].[dbo].[tbl_event_multi_contents] WHERE menuidx=" & menuidx & " and itemid in (" & Request.Form("cksel") & ")"
			'Response.write sqlStr
			'Response.end
			dbget.execute sqlStr
		IF Err.Number <> 0 THEN
			dbget.RollBackTrans 
			Call sbAlertMsg ("������ ó���� ������ �߻��Ͽ����ϴ�.", "back", "")
			response.End 
		END IF
	dbget.CommitTrans
	sqlStr = " Update [db_event].[dbo].[tbl_event_multi_contents_master]"
	sqlStr = sqlStr & " Set GroupItemType='T'"
	sqlStr = sqlStr & " ,GroupItemCheck='" & GroupItemCheck & "'"
	sqlStr = sqlStr & " Where idx=" & menuidx
	dbget.Execute sqlStr
	Response.write "<script>window.document.domain='10x10.co.kr';alert('���� �Ǿ����ϴ�.');parent.TnDelThemeItemBanner();</script>"
	response.End 
Else
	If vCnt >= 1 Then
	dbget.beginTrans

		For k=1 To vCnt
			sqlStr = " IF Not Exists(SELECT IDX FROM [db_event].[dbo].[tbl_event_multi_contents] WHERE grouptype='T' and itemid='" & arritemcode(k) & "' and menuidx="&menuidx& ")"			
			sqlStr = sqlStr + "	BEGIN "
			sqlStr = sqlStr+ " 			INSERT INTO [db_event].[dbo].[tbl_event_multi_contents] (menuidx, itemid, itemname, viewidx, grouptype)"
			sqlStr = sqlStr + "     	VALUES (" & menuidx & ", " & arritemcode(k) &",'" & arritemname(k) & "'," & k & ", 'T')"
			sqlStr = sqlStr + " 	END "
			sqlStr = sqlStr + " ELSE "
			sqlStr = sqlStr + " 	BEGIN "			
			sqlStr = sqlStr + "			UPDATE [db_event].[dbo].[tbl_event_multi_contents]"
			sqlStr = sqlStr + " 		SET viewidx ='" & k & "'"
			sqlStr = sqlStr + " 		WHERE grouptype='T' and menuidx = '" & menuidx & "' "
			sqlStr = sqlStr + " 		and itemid ="&arritemcode(k)&""
			sqlStr = sqlStr + " 	END "
			dbget.execute sqlStr
		IF Err.Number <> 0 THEN
			dbget.RollBackTrans 
			Call sbAlertMsg ("������ ó���� ������ �߻��Ͽ����ϴ�.", "back", "")
			response.End 
		END IF
		Next
		dbget.CommitTrans

		sqlStr = " Update [db_event].[dbo].[tbl_event_multi_contents_master]"
		sqlStr = sqlStr & " Set GroupItemType='T'"
		sqlStr = sqlStr & " ,GroupItemCheck='" & GroupItemCheck & "'"
		sqlStr = sqlStr & " Where idx=" & menuidx
		dbget.Execute sqlStr
	End If
End If

If upback = "Y" Then
	Response.write "<script>window.document.domain='10x10.co.kr';alert('��� �Ǿ����ϴ�.');parent.TnSaveThemeItemBanner();</script>"
End If
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->
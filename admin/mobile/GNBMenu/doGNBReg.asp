<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<%
'###############################################
' PageName : doGNBReg.asp
' Discription : GNBMEnu ó�� ������
' History : 2018.01.15 ������ ����
'###############################################

	Dim vMode, vIdx, vMenuCode, vStartDate, vStartDateSecond, vEndDate, vEndDateSecond
	Dim vMenuName, vLinkURL, vOrderBy, vIsNew, vIsUsing, sqlStr, vMenuPos
	

	vIdx = requestCheckVar(request("idx"), 10)
	vMode = requestCheckVar(request("mode"), 20)
	vMenuCode = requestCheckVar(request("MenuCode"), 30)
	vStartDate = requestCheckVar(request("StartDate"), 30)
	vStartDateSecond = requestCheckVar(request("StartDateSecond"), 30)
	vEndDate = requestCheckVar(request("EndDate"), 30)
	vEndDateSecond = requestCheckVar(request("EndDateSecond"), 30)
	vMenuName = requestCheckVar(request("MenuName"), 8)
	vLinkURL = requestCheckVar(request("LinkURL"), 8000)
	vOrderBy = requestCheckVar(request("OrderBy"), 30)
	vIsNew = requestCheckVar(request("IsNew"), 30)
	vIsUsing = requestCheckVar(request("IsUsing"), 30)
	vMenuPos = requestCheckVar(request("menupos"), 50)

'// ��忡 ���� �б�
Select Case vMode
	Case "add"
		'�ű� ���
		sqlStr = "Insert Into db_sitemaster.[dbo].[tbl_GNBMenuManagement] " &_
					" (MenuCode, MenuName , LinkURL , StartDate , EndDate, RegDate, LastUpDate, AdminId, LastAdminId, OrderBy, IsNew, IsUsing ) values " &_
					" ('" & vMenuCode &"'" &_
					" ,'" & vMenuName &"'" &_
					" ,'" & vLinkURL &"'" &_
					" ,'" & vStartDate&" "&vStartDateSecond&"'" &_
					" ,'" & vEndDate&" "&vEndDateSecond&"'" &_
					" ,getdate()" &_
					" ,getdate()" &_
					" ,'" & session("ssBctId") &"'" &_
					" ,'" & session("ssBctId") &"'" &_
					" ,'" & vOrderBy &"'" &_
					" ,"&vIsNew &_
					" ,"&vIsUsing &_
					")"
		dbget.Execute(sqlStr)

	Case "modi"
		'���� ����
		sqlStr = " UPDATE db_sitemaster.[dbo].[tbl_GNBMenuManagement] SET " &_
					" MenuName = '"&vMenuName&"'" &_
					" ,LinkURL = '"&vLinkURL&"'" &_
					" ,StartDate = '"&vStartDate&" "&vStartDateSecond&"'" &_
					" ,EndDate = '"&vEndDate&" "&vEndDateSecond&"'" &_
					" ,LastUpdate = getdate()" &_
					" ,LastAdminId = '"&session("ssBctId")&"'" &_
					" ,OrderBy = '"&vOrderBy&"'" &_
					" ,IsNew = "&vIsNew &_
					" ,IsUsing = "&vIsUsing &_
					" where idx= "&vIdx
		dbget.Execute(sqlStr)

	Case "modiAll"
		'�ش� MenuCode ��ü ������ ó��
		sqlStr = " UPDATE db_sitemaster.[dbo].[tbl_GNBMenuManagement] SET " &_
					" IsUsing = 0 "&_
					" ,LastUpdate = getdate()" &_
					" ,LastAdminId = '"&session("ssBctId")&"'" &_
					" where MenuCode= '"&vMenuCode&"' "
		dbget.Execute(sqlStr)

End Select

%>

<script>
<!--
	// ������� ����
	<% if trim(vMode)="modiAll" then %>
		alert("������ ó���Ǿ����ϴ�.");
		self.location = "index.asp?menupos=<%=vmenupos%>&MenuCode=<%=vMenuCode%>";
	<% else %>
		alert("�����߽��ϴ�.");
		window.opener.document.location.href = window.opener.document.URL;    // �θ�â ���ΰ�ħ
	<% end if %>

	 self.close();        // �˾�â �ݱ�
//-->
</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->

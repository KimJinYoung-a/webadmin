<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  �������� �̺�Ʈ ���ڵ� ��ǰ�߰�
' History : 2012.04.24 ���ر� ����
'			2012.08.09 �ѿ�� ����
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<%
Dim mode, referer, itemid ,evt_code ,addSql ,sqlStr , itemoption , itemgubun
Dim itemidarr, itemoptionarr ,itemgubunarr , i
	mode 			= requestCheckVar(Request("mode"),32)
	itemid 			= requestCheckVar(Request("itemid"),10)
	itemoption 		= requestCheckVar(Request("itemoption"),4)
	itemgubun 		= requestCheckVar(Request("itemgubun"),2)
	itemidarr 		= Trim(Request("itemidarr"))
	itemoptionarr 	= Request("itemoptionarr")
	itemgubunarr 	= Request("itemgubunarr")
	evt_code 		= requestCheckVar(request("evt_code"),10)
	referer 		= request.ServerVariables("HTTP_REFERER")

If itemidarr = "" Then
	Response.Write "<script type='text/javascript'>alert('���ڰ��� �����ϴ�');</script>"
	response.end	:	dbget.close()
End If
if len(itemidarr) < 11 then
	response.write "<script type='text/javascript'>"
	response.write "	alert('���ڵ��� ���̰� ª���ϴ�.\n�����ڵ峪 ������ڵ带 �ٽ� Ȯ����, �Է� �ϼ���.');"
	response.write "</script>"
	response.end	:	dbget.close()
end if

Dim vTempArr, vItemGubun, vShopItemID, vItemOption, vIsOK, vQuery, vCount
	vTempArr		= itemidarr
	vItemGubun		= Trim(Left(vTempArr, 2))
	vItemOption		= Trim(Right(vTempArr, 4))
	vTempArr		= Trim(Right(vTempArr,Len(vTempArr)-2))
	vTempArr		= Trim(Left(vTempArr,Len(vTempArr)-4))
	vShopItemID		= Trim(vTempArr) ''Trim(Format00(6,vTempArr))
	vIsOK			= "x"
	vCount			= 0

If mode = "itemadd" Then

	if IsNumeric(vShopItemID) = TRUE and ((len(itemidarr) = 12) or (len(itemidarr) = 14)) then

		vQuery = "SELECT COUNT(shopitemid) FROM [db_shop].[dbo].[tbl_shop_item]"
		vQuery = vQuery & " WHERE itemgubun = '" & vItemGubun & "'"
		vQuery = vQuery & " AND shopitemid = '" & vShopItemID & "' AND itemoption = '" & vItemOption & "'"

		'response.write vQuery & "<Br>"
		rsget.Open vQuery,dbget,1

		If Not rsget.Eof Then
			vCount = rsget(0)
		End If

		rsget.close()
	end if

	If vCount > 0 Then
		vIsOK = "o"
	Else
		vQuery = "SELECT COUNT(shopitemid) FROM [db_shop].[dbo].[tbl_shop_item] WHERE extbarcode = '" & itemidarr & "'"

		'response.write vQuery & "<Br>"
		rsget.Open vQuery,dbget,1

		If Not rsget.Eof Then
			If rsget(0) > 0 Then
				vIsOK = "o"
			End If
		End If

		rsget.close()
	End IF
End If


If vIsOK = "o" Then
%>
	<form name="frm" method="post" action="/admin/offshop/event_off/eventitem_off_process.asp">
		<input type="hidden" name="evt_code" value="<%=evt_code%>">
		<input type="hidden" name="mode" value="<%=mode%>">
		<input type="hidden" name="itemidarr" value="<%=vShopItemID%>,">
		<input type="hidden" name="itemoptionarr" value="<%=vItemOption%>,">
		<input type="hidden" name="itemgubunarr" value="<%=vItemGubun%>,">
	</form>

	<script language="javascript">
		document.frm.submit();
	</script>
<%
Else

	Response.Write "<script type='text/javascript'>alert('�������� �ʴ� ��ǰ�Դϴ�.');</script>"
	response.end	:	dbget.close()
End If
%>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/giftManager/GiftManagerCls.asp"-->

<%

dim cdL, cdM, cdS

cdL= request("cdL")
cdM= request("cdM")
cdS= request("cdS")



function defbgColor(byVal v ,byVal c)
	if  v <> c then
		defbgColor="#FFFFFF"
	else
		defbgColor="#D2E1FF"
	end if
end function


%>
<script language="javascript" type="text/javascript">
function PopAddMenu(dep,cdl,cdm,cds){
	window.open('Pop_Menu_Add.asp?Depth=' + dep + '&cdl=' + cdl + '&cdm=' + cdm + '&cds=' +cds,'pop','width=500,height=300');
}
function PopEditMenu(dep,cdl,cdm,cds){
	window.open('Pop_Menu_Edit.asp?Depth=' + dep + '&cdl=' + cdl + '&cdm=' + cdm + '&cds=' +cds,'pop','width=500,height=300');
}
function DelMenu(dep,cdl,cdm,cds){
	var ret = confirm('���� �Ͻðڽ��ϱ�?');
	if (ret){
		window.open('Menu_Process.asp?mode=del&Depth=' + dep + '&LCode=' + cdl + '&MCode=' + cdm + '&SCode=' +cds,'pop','width=50,height=50');
	}
}
function CreateMenu(){
	window.open('<%= wwwUrl %>/chtml/make_giftManager_Menu.asp','pop','width=500,height=300');
}
</script>
<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
<tr>
	<td valign="top" width="220">
		<table border="0" cellpadding="2" cellspacing="1" class="a">
			<tr>
				<td><input type="button" value="�޴� ����" class="button" onclick="CreateMenu();"></td>
			</tr>
		</table>
		<table border="0" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
			<tr bgcolor="<%= adminColor("tabletop") %>">
				<td align="center" width="25">����</td>
				<td align="center" width="25">�ڵ�</td>
				<td align="center" width="80">�� ī�׸�</td>
				<td align="center" width="85">����</td>
			</tr>
			<%
			'// �� ī�װ� ���
			dim objMnL,objMnM,objMnS ,i
			set objMnL = new giftManagerMenu
			objMnL.getMenuListLarge 
			%>
			<% if objMnL.FResultCount<>0 then %>
			<% for i = 0 to objMnL.FResultCount -1 %>
			<tr bgcolor="<%= defbgColor(cdL,objMnL.FItemList(i).LCode) %>">
				<td align="center"><%= objMnL.FItemList(i).OrderNo %></td>
				<td align="center"><%= objMnL.FItemList(i).LCode %></td>
				<td align="center"><a href="?cdL=<%= objMnL.FItemList(i).LCode %>"><%= objMnL.FItemList(i).LCodeNm %></a></td>
				<td align="center">
					<input type="button" value="����" class="button" onclick="PopEditMenu('L','<%= objMnL.FItemList(i).LCode %>','','');">
					<input type="button" class="button" value="����" onclick="DelMenu('L','<%= objMnL.FItemList(i).LCode %>','','');">
				</td>	
			</tr>
			<% next %>
			<% end if %>
			<tr bgcolor="#FFFFFF">
				<td align="center" colspan="4"><input type="button" class="button" value="�߰�" onclick="PopAddMenu('L','','','');"></td>
			</tr>
		</table>
	
		<%
		' // �� ī�װ� ��� 
		if cdL <>"" then
		set objMnM = new giftManagerMenu
		objMnM.FRectCDL = cdL
		objMnM.getMenuListMid
		%>
		<table border="0" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
			
			<tr bgcolor="<%= adminColor("tabletop") %>">
				<td align="center" width="25">����</td>
				<td align="center" width="25">�ڵ�</td>
				<td align="center" width="80">�� ī�׸�</td>
				<td align="center" width="85">����</td>
			</tr>
			<% if objMnM.FResultCount<>0 then %>
			<% for i = 0 to objMnM.FResultCount -1 %>
			<tr bgcolor="<%= defbgColor(cdM,objMnM.FItemList(i).MCode) %>">
				<td align="center"><%= objMnM.FItemList(i).OrderNo %></td>
				<td align="center"><%= objMnM.FItemList(i).MCode %></td>
				<td align="center"><a href="?cdL=<%= objMnM.FItemList(i).LCode %>&cdM=<%= objMnM.FItemList(i).MCode %>"><%= objMnM.FItemList(i).MCodeNm %></a></td>
				<td align="center">
					<input type="button" value="����" class="button" onclick="PopEditMenu('M','<%= objMnM.FItemList(i).LCode %>','<%= objMnM.FItemList(i).MCode %>','');">
					<input type="button" class="button" value="����" onclick="DelMenu('M','<%= objMnM.FItemList(i).LCode %>','<%= objMnM.FItemList(i).MCode %>','');">
				</td>
			</tr>
			<% next %>
			<% end if %>
			<tr bgcolor="#FFFFFF">
				<td align="center" colspan="4"><input type="button" class="button" value="�߰�" onclick="PopAddMenu('M','<%= cdL %>','','');"></td>
			</tr>
		</table>
		<% end if %>
		<%
		'// �� ī�װ� ��� 
		if cdM <>"" then
			set objMnS = new giftManagerMenu
			objMnS.FRectCDL = cdL
			objMnS.FRectCDM = cdM
			objMnS.getMenuListSmall
		%>
		<table border="0" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
			<tr bgcolor="<%= adminColor("tabletop") %>">
				<td align="center" width="25">����</td>
				<td align="center" width="25">�ڵ�</td>
				<td align="center" width="80">�� ī�׸�</td>
				<td align="center" width="85">����</td>
			</tr>
			<% if objMnS.FResultCount<>0 then %>
			<% for i = 0 to objMnS.FResultCount -1 %>
			<tr bgcolor="<%= defbgColor(cdS,objMnS.FItemList(i).SCode) %>">
				<td align="center"><%= objMnS.FItemList(i).OrderNo %></td>
				<td align="center"><%= objMnS.FItemList(i).SCode %></td>
				<td align="center"><a href="?cdL=<%= objMnS.FItemList(i).LCode %>&cdM=<%= objMnS.FItemList(i).MCode %>&cdS=<%= objMnS.FItemList(i).SCode %>"><%= objMnS.FItemList(i).SCodeNm %></a></td>
				<td align="center">
					<input type="button" value="����" class="button" onclick="PopEditMenu('S','<%= objMnS.FItemList(i).LCode %>','<%= objMnS.FItemList(i).MCode %>','<%= objMnS.FItemList(i).SCode %>');">
					<input type="button" class="button" value="����" onclick="DelMenu('S','<%= objMnS.FItemList(i).LCode %>','<%= objMnS.FItemList(i).MCode %>','<%= objMnS.FItemList(i).SCode %>');"
				</td>
			</tr>
			<% next %>
			<% end if %>
			<tr bgcolor="#FFFFFF">
				<td align="center" colspan="4"><input type="button" class="button" value="�߰�" onclick="PopAddMenu('S','<%= cdL %>','<%= cdM %>','');"></td>
			</tr>
		</table>
		<% end if %>
	</td>
	<td style="padding-left:5;" valign="top">
		<%' if cdS<>"" then %>
		<iframe src="iframe_itemList.asp?cdL=<%= cdL %>&cdM=<%= cdM %>&cdS=<%= cdS %>"  frameborder="0" width="820" height="750"></iframe>
		<%' end if %>
	</td>
</tr>
</table>

<% 
set objMnL = nothing 
set objMnM = nothing 
set objMnS = nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
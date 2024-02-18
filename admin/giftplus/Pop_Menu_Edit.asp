<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  ����Ʈ�÷���
' History : 2010.04.02 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/classes/giftplus/giftplus_cls.asp"-->

<%
dim Depth, cdL, cdM, cdS, SortMethod, MnL_viewType ,objView ,ECodeNm , EOrderNo ,imageName,imageMain
dim listtype
	Depth = request("depth")
	cdL= request("cdL")
	cdM= request("cdM")
	cdS= request("cdS")

set objView = new giftManagerView
	objView.getMenuView cdL,cdM,cdS

	if SortMethod ="" then SortMethod ="cashHigh"

	if Depth = "L" then		
		ECodeNm 	= objView.LCodeNm
		EOrderNo 	= objView.OrderNo
		
	elseif Depth = "M" then		
		ECodeNm 	= objView.MCodeNm
		EOrderNo 	= objView.OrderNo
	elseif Depth = "S" then				
		ECodeNm 	= objView.SCodeNm
		EOrderNo 	= objView.OrderNo
	end if

	if Depth = "L" then	
		imageName="Menu" & cdL 
	elseif Depth = "M" then	
		imageName="MidTop" & cdL & "-" & cdM 
	elseif Depth = "S" then	
		imageName="Mania" & cdL & "-" & cdM & "-" & cdS 
		imageMain="ManiaMain" & cdL & "-" & cdM & "-" & cdS 
	end if		

	if cdL <> "" then
	'listtype = objView.listtype
	listtype = getlisttype(cdl)
	end if
	
	if listtype = "" then listtype = "menu"

if listtype = "search" and DEPTH = "S" then		
	response.write "<script>alert('ǥ�������� �˻����ϰ�� ��ī�װ��� �߰��ϽǼ� �����ϴ�'); self.close();</script>"
	dbget.close() : response.end
end if
%>

<script language="javascript">

function subchk(){
	//if (isNaN(UpdateFRM.viewidx.value)) {
	//	alert('���ڸ� �Է°����մϴ�');
	//	return false;
	//}
}

function popInputImg(ImgNm,Frmvalue){
	var pop = window.open('Pop_Menu_InputImage.asp?imageName=' + ImgNm +'&Frmvalue=' + Frmvalue ,'pp','width=300,height=150,resizable=yes');
}

function searchCashEdit(cdL,cdM,cdS){
	var pop = window.open('Pop_Menu_CashEdit.asp?cdL=' + cdL + '&cdM=' + cdM + '&cdS=' + cdS,'pp','width=500,height=300,status=yes,resizable=yes');
}

document.domain = "10x10.co.kr";
window.resizeTo(430,400);

</script>

<table width="400" border="0" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="UpdateFRM" action="Menu_Process.asp" target="" onSubmit="return subchk();">
<input type="hidden" name="mode" value="edit">
<input type="hidden" name="Depth" value="<%= Depth %>">
<input type="hidden" name="LCode" size="4" value="<%= objView.LCode %>" />
<input type="hidden" name="MCode" size="4" value="<%= objView.MCode %>" />
<input type="hidden" name="SCode" size="4" value="<%= objView.SCode %>" />
<tr>
	<td width="130" bgcolor="#FFFFFF"></td>
	<td bgcolor="#FFFFFF"></td>
</tr>

<% 
IF objView.LCode <>"" then 
%>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>" align="center">�� ī�װ�</td>
	<td bgcolor="#FFFFFF"> [<font color="red"><%= objView.LCode %></font>] <%= objView.LCodeNm %>
</tr>
<% END IF %>
<% IF objView.MCode <>"" then %>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>" align="center">�� ī�װ�</td>
	<td bgcolor="#FFFFFF"> [<font color="red"><%= objView.MCode %></font>] <%= objView.MCodeNm %>
</tr>
<% END IF %>
<% IF objView.SCode <>"" then %>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>" align="center">�� ī�װ�</td>
	<td bgcolor="#FFFFFF"> [<font color="red"><%= objView.SCode %></font>] <%= objView.SCodeNm %>
</tr>
<% END IF %>
<% 
'����Ʈ�� ���� �޴��� ���� ����� ��ī�װ� ������ �ƴҰ�쿡�� ���� ���� �����ϰ���
if listtype = "menu" and DEPTH <> "L" then 
%>	
<tr>
	<td colspan="2" height="20" bgcolor="#FFFFFF" align="right"><input type="button" class="button" value="�˻����ݰ���" onclick="searchCashEdit('<%= objView.LCode %>','<%= objView.MCode %>','<%= objView.SCode %>')"></td>
</tr>
<% end if %>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>" align="center">����</td>
	<td bgcolor="#FFFFFF"><input type="text" size="4" name="OrderNo" value="<%= EOrderNo %>">(0 ~ 99)</td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>" align="center">ī�װ���</td>
	<td bgcolor="#FFFFFF"><input type="text" size="16" name="CodeNm" value="<%= ECodeNm %>"></td>
</tr>

<!-- Depth �� ���� -->
<% SELECT CASE DEPTH %>
<% CASE "L" %>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>" align="center"  align="center">�޴� �̹���<br>(Ȱ��)</td>
	<td bgcolor="#FFFFFF">
		<input type="text" name="LCodeImgON" id="LCodeImgON" value="<%= objView.LCodeImgON %>" size="25" />
		<input type="button" class="button" value="�̹����ֱ�" onclick="popInputImg('<%= imageName %>on','LCodeImgON');">
	</td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>" align="center"  align="center">�޴� �̹���<br>(��Ȱ��)</td>
	<td bgcolor="#FFFFFF">
		<input type="text" name="LCodeImgOFF" id="LCodeImgOFF" value="<%= objView.LCodeImgOFF %>" size="25" />
		<input type="button" class="button" value="�̹����ֱ�" onclick="popInputImg('<%= imageName %>','LCodeImgOFF');">
	</td>
</tr>
<% CASE "M" %>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>" align="center">��� �̹���</td>
	<td bgcolor="#FFFFFF">
		<input type="text" name="MCodeTopImg" id="MCodeTopImg" value="<%= objView.MCodeTopImg %>" size="25" />
		<input type="button" class="button" value="�̹����ֱ�" onclick="popInputImg('<%= imageName %>','MCodeTopImg');">
	</td>
</tr>
<% END SELECT %>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>" align="center">��뿩��</td>
	<td bgcolor="#FFFFFF">		
		<input type="radio" name="isusing" value="Y" <% IF objView.IsUsing="Y" Then response.write "checked" %>> ��� 
		<input type="radio" name="isusing" value="N" <% IF objView.IsUsing="N" Then response.write "checked" %>>������</td>
</tr>
<% IF DEPTH = "L" THEN %>
<tr>
<td bgcolor="<%= adminColor("tabletop") %>" align="center">ǥ������</td>
<td bgcolor="#FFFFFF">
	<% drawListType "listtype" , listtype, "" %>
</td>
<% END IF %>	
<tr>
	<td bgcolor="#FFFFFF" colspan="2" align="center"><input type="submit" class="button" value="����"></td>
</tr>
</form>
</table>

<% 
set objView = nothing 
%>

<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
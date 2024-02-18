<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description : ����
' History : ������ ����
'			2017.04.13 �ѿ�� ����(���Ȱ���ó��)
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/offshop/dishcorrectcls.asp"-->
<%
dim makerid, shopid, availbojung, research
	makerid = requestCheckVar(request("makerid"),32)
	shopid = requestCheckVar(request("shopid"),32)
	availbojung = requestCheckVar(request("availbojung"),2)
	research = requestCheckVar(request("research"),2)

if (research="") and (availbojung="") then availbojung="on"

dim offbojung
set offbojung = new COffShopDishCorrect
offbojung.FRectShopId = shopid
offbojung.FRectMakerid = makerid
offbojung.FRecAvailbojung = availbojung

if (makerid<>"") and (shopid<>"") then
	offbojung.GetDishValidList
end if

dim i, bojungno
%>

<table width="800" border="0" cellpadding="5" cellspacing="0" bgcolor="#CCCCCC">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="research" value="on">
	<tr>
		<td class="a" >
			�� : <% drawSelectBoxOffShop "shopid",shopid %> &nbsp;&nbsp;
			��ü:<% drawSelectBoxDesignerwithName "makerid",makerid  %> &nbsp;&nbsp;
			<input type=checkbox name="availbojung" <% if availbojung="on" then response.write "checked" %> >��ȿ�������˻�
		</td>
		<td class="a" align="right">
			<a href="javascript:document.frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
		</td>
	</tr>
	</form>
</table>
<br>

<b>[OFF]����_�������>>[���]���������Ȳ ���� ������� Ȯ�� ���� : ���Ͻð� �۾����ּ���</b>
<table width="800" cellspacing="1" class="a" bgcolor=#3d3d3d>
<tr bgcolor="#DDDDFF" align=center>
    <td width="50">�̹���</td>
	<td width="80">���ڵ�</td>
	<td width="100">��ǰ��</td>
	<td width="60">�ɼǸ�</td>
	<td width="50">�¶���<br>�ǸŰ�</td>
	<td width="50">����<br>�ǸŰ�</td>
	<td width="50">������<br>�ǸŰ�</td>
	<td width="30">������</td>
	<td width="30">����y��<br>����</td>
	<td width="50">�����Ǹ�</td>
	<td width="30">����</td>
	<td width="30">�ý���<br>���</td>
	<td width="30">������</td>
</tr>
<% for i=0 to offbojung.FresultCount-1 %>
<tr bgcolor="#FFFFFF">
	<td><img src="<%= offbojung.FItemList(i).Fimgsmall %>" onError="this.src='http://webimage.10x10.co.kr/images/no_image.gif'" width=50 height=50></td>
	<td><%= offbojung.FItemList(i).GetBarCode %></td>
	<td><%= offbojung.FItemList(i).FItemName %></td>
	<td><%= offbojung.FItemList(i).FItemOptionName %></td>
	<td align=right><%= offbojung.FItemList(i).Fonlinesellcash %></td>
	<td align=right><%= offbojung.FItemList(i).Fofflinesellcash %></td>
	<td align=right><%= offbojung.FItemList(i).Fipchulsellcash %></td>
	<td align=center><%= offbojung.FItemList(i).Fipchulno %></td>
	<td align=center><%= offbojung.FItemList(i).Fipchulsamesellno %></td>
	<td align=right><%= offbojung.FItemList(i).Fipchuldiffsellprice %></td>
	<td align=center><%= offbojung.FItemList(i).Fipchuldiffsellno %></td>
	<td align=center><%= offbojung.FItemList(i).Fstockcurrno %></td>
	<% if (offbojung.FItemList(i).Fonlinesellcash<>offbojung.FItemList(i).Fipchulsellcash) or (offbojung.FItemList(i).Fofflinesellcash<>offbojung.FItemList(i).Fipchulsellcash) then %>
	<td align=center><b><%= offbojung.FItemList(i).GetMayBojungCount   %></b></td>
	<% else %>
	<td align=center><font color="#CCCCCC"><%= offbojung.FItemList(i).GetMayBojungCount %></font></td>
	<% end if %>
</tr>
<% next %>
</table>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
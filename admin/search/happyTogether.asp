<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbCTopen.asp" -->
<!-- #include virtual="/lib/db/dbAnalopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/happyTogetherCls.asp" -->
<!-- #include virtual="/lib/classes/items/new_itemcls.asp"-->
<%

dim itemid, cnt, pcnt, ordby, samemaker, simCate
dim i
dim research

research 	= request("research")
itemid 		= Trim(request("itemid"))
cnt 		= Trim(request("cnt"))
pcnt 		= Trim(request("pcnt"))
ordby 		= requestCheckVar(request("ordby"),2)
samemaker 		= Trim(request("samemaker"))
simCate 		= Trim(request("simCate"))

if (pcnt = "") then
	pcnt = 0.1
end if

if (cnt = "") then
	cnt = 1
end if

if (ordby = "") then
	ordby = "tc"
end if

if (simCate = "") and (research = "") then
	simCate = "Y"
end if

if (samemaker = "") and (research = "") then
	samemaker = "Y"
end if

dim oitem
set oitem = new CItemInfo
	oitem.FRectItemID = itemid

	if itemid<>"" then
		oitem.GetOneItemInfo
	end if


'// ============================================================================
dim oHappyTogether

set oHappyTogether = new CHappyTogether

oHappyTogether.FRectItemID			= itemid
oHappyTogether.FRecCnt				= cnt
oHappyTogether.FRecPCnt				= pcnt
oHappyTogether.FRecOrderBy			= ordby

oHappyTogether.FRecSimCateOnly		= simCate
oHappyTogether.FRecSameUpcheOnly	= samemaker

if itemid<>"" then
	''oHappyTogether.GetHappyTogetherRawList
	oHappyTogether.GetHappyTogetherBuyAlsoList
end if

%>

<script language='javascript'>

function jsSearch(itemid) {
	var frm = document.frm;
	frm.itemid.value = itemid;
	frm.submit();
}

function jsOrderType(otp) {
	var frm = document.frm;
	frm.ordby.value = otp;
	frm.submit();
}

</script>

<!-- �˻� ���� -->
<form name="frm" method="get" action="">
<input type="hidden" name="research" value="on">
<input type="hidden" name="page" value="1">
<input type="hidden" name="ordby" value="<%=ordby%>">
<input type="hidden" name="menupos" value="<%= request("menupos") %>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="1" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
		<td align="left" height="30" >
			��ǰ�ڵ� : <input type="text" class="text" name="itemid" value="<%= itemid %>">
		</td>
		<td rowspan="1" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button" value="�˻�" onClick="javascript:document.frm.submit();">
		</td>
	</tr>
	<!--
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left" height="30" >
			<input type="checkbox" name="samemaker" value="Y" <% if (samemaker = "Y") then %>checked<% end if %> > ������� ���ɻ�ǰ��
			<input type="checkbox" name="simCate" value="Y" <% if (simCate = "Y") then %>checked<% end if %> > ���� ī�װ� ��ǰ��
		</td>
	</tr>
	-->
</table>
</form>
<!-- �˻� �� -->

<p>

	<!--
		 * ���� ī�װ�<br>
		 �����ι��� : ���� | ������/�ڵ��� : ���� | ķ��/Ʈ���� : ���� | ���� : ����<br>
		 ���� : Ȩ���׸��� | Ȩ���׸��� : ���� | Űģ/Ǫ�� : ���� | �м��Ƿ� : ����<br>
		 ����/����/�־� : �Ƿ� | ��Ƽ/���̾�Ʈ : ���� | ���̺�/Ű�� : ���� | Ĺ&���� : ����
	   -->
	* ����ǰ�� �����մϴ�. �Ͻ�ǰ�� ��ǰ�� �����մϴ�.
<p>

<% if (oitem.FResultCount>0) then %>
	<table width="100%" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr bgcolor="#FFFFFF">
		<td rowspan=6 width="110" valign=top align=center><img src="<%= oitem.FOneItem.FListImage %>" width="100" height="100"></td>
		<td width="60">��ǰ�ڵ�</td>
		<td width="300">
			<%= oitem.FOneItem.FItemID %>
		</td>
		<td colspan="5">
		</td>

	</tr>

	<tr bgcolor="#FFFFFF">
		<td>�귣��ID</td>
		<td><%= oitem.FOneItem.FMakerid %></td>
		<td>�Ǹſ���</td>
		<td colspan=4><font color="<%= ynColor(oitem.FOneItem.FSellyn) %>"><%= oitem.FOneItem.FSellyn %></font></td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td>��ǰ��</td>
		<td><a href="http://www.10x10.co.kr/shopping/category_prd.asp?itemid=<%= oitem.FOneItem.FItemID %>" target="_blank"><%= oitem.FOneItem.FItemName %></a></td>
		<td>��뿩��</td>
		<td colspan=4><font color="<%= ynColor(oitem.FOneItem.FIsUsing) %>"><%= oitem.FOneItem.FIsUsing %></font></td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td>�ǸŰ�</td>
		<td>
			<%= FormatNumber(oitem.FOneItem.FSellcash,0) %>
		</td>
		<td>��������</td>
		<td colspan="4">
			<%= fncolor(oitem.FOneItem.Fdanjongyn,"dj") %>
			<% if oitem.FOneItem.Fdanjongyn="N" then %>
			������
			<% end if %>
		</td>
	</tr>

	</table>
<% end if %>

<p>

	<% if (oHappyTogether.FTotalCount > 0) then %>
	�˻���� : <%= oHappyTogether.FTotalCount %>
<p>
	<% end if %>

<!-- ����Ʈ ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
 	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    	<td width="65" height="30" onclick="jsOrderType('ia');" style="cursor:pointer;<%=chkIIF(ordby="ia","font-weight:bold;","")%>">itemid A</td>
		<td width="65" onclick="jsOrderType('ib');" style="cursor:pointer;<%=chkIIF(ordby="ib","font-weight:bold;","")%>">itemid B</td>
		<td width="105">�̹���</td>
		<td width="120">�귣��</td>
		<td width="400">��ǰ��</td>
		<td width="110" onclick="jsOrderType('tc');" style="cursor:pointer;<%=chkIIF(ordby="tc","font-weight:bold;","")%>">��ȸ�Ǽ�<br />(�ֱ�3��)</td>
		<td width="110" onclick="jsOrderType('oc');" style="cursor:pointer;<%=chkIIF(ordby="oc","font-weight:bold;","")%>">���ŰǼ�<br />(�ֱ�2��)</td>
		<td width="150"></td>
		<td>���</td>
	</tr>
	<%
	for i = 0 To oHappyTogether.FTotalCount - 1
	%>
	<tr align="center" bgcolor="#FFFFFF">
		<td align="center" height="30">
			<%= oHappyTogether.FItemList(i).FitemidA %>
		</td>
		<td align="center"><a href="javascript:jsSearch(<%= oHappyTogether.FItemList(i).FitemidB %>)"><%= oHappyTogether.FItemList(i).FitemidB %></a></td>
		<td align="center"><img src="<%= oHappyTogether.FItemList(i).Flistimage %>"></td>
		<td align="left"><%= oHappyTogether.FItemList(i).Fmakerid %></td>
		<td align="left"><a href="http://www.10x10.co.kr/shopping/category_prd.asp?itemid=<%= oHappyTogether.FItemList(i).FitemidB %>" target="_blank"><%= oHappyTogether.FItemList(i).Fitemname %></a></td>
		<td align="center"><%= oHappyTogether.FItemList(i).FtotCnt %></td>
		<td align="center"><%= oHappyTogether.FItemList(i).Fcnt %></td>
		<td align="center"></td>

		<td align="left">

		</td>
	</tr>
	<%
	next
	%>
	<% if (oHappyTogether.FTotalCount = 0) then %>
	<tr align="center" bgcolor="#FFFFFF">
		<td height="30" colspan="10">
			�˻������ �����ϴ�.
		</td>
	</tr>
	<% end if %>
</table>
<%
set oHappyTogether = Nothing
%>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbCTclose.asp" -->
<!-- #include virtual="/lib/db/dbAnalclose.asp" -->

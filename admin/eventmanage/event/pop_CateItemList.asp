<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/items/itemcls_2008.asp"-->
<%
'###############################################
' PageName : pop_CateItemList.asp
' Discription : ī�װ��� ��ǰ ��� �˾�
'	- ���� ī�װ��� ��ǰ�� �����ְ� ������ ��ǰ�ڵ带 ��ȯ
' History : 2008.04.01 ������ : ����
'			2008.06.11 ������ ���� : ���Ĺ�� ����
'###############################################

dim makerid,itemid,itemname
dim page,sellyn,packyn
dim target
dim cdl,cdm,cds
dim sortdiv

cdl = request("cdl")
cdm = request("cdm")
cds = request("cds")

makerid = request("makerid")
itemid = request("itemid")
itemname = request("itemname")
sellyn = request("sellyn")
page = request("page")
target= request("target")
sortdiv =request("sd")
if page="" then page=1

dim oItem
set oItem = new CItem
oItem.FCurrPage = page
oItem.FPageSize = 20
oItem.FRectItemName = itemname
oItem.FRectMakerid = makerid
oItem.FRectItemid = itemid
oItem.FRectSellYN = sellyn
oItem.FRectCate_Large = cdl
oItem.FRectCate_Mid = cdm
oItem.FRectCate_Small = cds
oItem.FRectSortDiv = sortdiv
oItem.GetItemList

dim i
%>
<script>
function SelectItems(){
	var tg = eval('opener.frmWin.itemuse');
	var tgvalue="";
	var tgid = eval('opener.frmWin.itemuse_itemid');
	var tgidvalue="";

	var frm;
	var pass = false;
	var temp = 0;

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			pass = ((pass)||(frm.cksel.checked));
		}
	}

	if (!pass) {
		alert('���� ��ǰ�� �����ϴ�.');
		return;
	}

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			if (frm.cksel.checked){
				tgvalue = tgvalue + "[" + frm.itemid.value + "]" + frm.itemname.value + " " ;
				tgidvalue = tgidvalue + frm.itemid.value + " " ;
				temp = temp + 1;
			}
		}
	}

	if (temp > 1)
	{
		alert("�Ѱ��� ������ �����մϴ�.");
		return;
	}

	tg.focus();
	tg.value = tgvalue;
	tgid.value = tgidvalue;
	opener.ViewByteLength();
	window.close();

}
</script>
<body bgcolor="#F4F4F4">
<!-- �ش� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td>
		<table width="100%" align="center" cellpadding="1" cellspacing="0" class="a" bgcolor="#999999">
		<tr>
			<td width="400" style="padding:5; border-top:1px solid #999999;border-left:1px solid #999999;border-right:1px solid #999999"  background="/images/menubar_1px.gif">
				<font color="#333333"><b>ī�װ� ��ǰ ����/�߰�</b></font>
			</td>
			<td align="right" style="border-bottom:1px solid #999999" bgcolor="#F4F4F4">&nbsp;</td>
		</tr>
		</table>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td style="padding:5; border-bottom:1px solid #999999;border-left:1px solid #999999;border-right:1px solid #999999" bgcolor="#FFFFFF">
		ī�װ��� ���� ��ǰ�� �˻�/�����մϴ�.
	</td>
</tr>
<tr><td height="10"></td></tr>
</table>
<!-- �˻� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="page" value="1">
<input type="hidden" name="sd" value="<%=sortdiv%>">
<input type="hidden" name="target" value="<%= target %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
		<table width="100%" border="0" cellpadding="0" cellspacing="0" class="a">
		<tr>
			<td>
			�����̳� :
			<% drawSelectBoxDesigner "makerid",makerid %>
			��ǰID :
			<input type="text" class="text" name="itemid" value="<%= itemid %>" size="10" maxlength="9">
			��ǰ�� :
			<input type="text" class="text" name="itemname" value="<%= itemname %>" size="8" maxlength="32">
			�Ǹſ��� :
			<select name="sellyn" class="select">
		                 	<option value='' selected>����</option>
		                 	<option value='Y' <% if sellyn="Y" then response.write "selected" %> >Y</option>
		                 	<option value='N' <% if sellyn="N" then response.write "selected" %> >N</option>
		         	</select>
			</td>
		</tr>
		<tr>
			<td>
				<!-- #include virtual="/common/module/categoryselectbox.asp"-->
			</td>
		</tr>
		</table>
	</td>
	<td width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="submit" class="button_s" value="�˻�">
	</td>
</tr>
</form>
</table>
<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
<form name="frmttl" onsubmit="return false;">
<tr  style="padding:10 0 10 0;">
	<td>
		<!--<input type="button" class="button" value="��ü����" onClick="AnSelectAllFrame(true)">//-->
	</td>
	<td align="right">
		<input type="button" class="button" value="��ǰ����" onClick="SelectItems()">
	</td>
</tr>
</form>
</table>
<!-- ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="6">
		�˻���� : <b><%= FormatNumber(oItem.FTotalCount,0) %></b>
		&nbsp;
		������ : <b><%= FormatNumber(page,0) %> / <%= FormatNumber(oItem.FTotalPage,0) %></b>
		&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
		<font color="red"><b>�� �ݵ�� ��ǰ�� �Ѱ��� �����ϼ���.</b></font>
	</td>
	<td colspan="2" align="right">����:
		<select name="sortDiv" onchange="frm.sd.value = this.value; frm.submit();">
		<option value="new" <% IF sortDiv="new" Then response.write "selected" %> >�Ż�ǰ��</option>
		<option value="cashH" <% IF sortDiv="cashH" Then response.write "selected" %>>�������ݼ�</option>
		<option value="cashL" <% IF sortDiv="cashL" Then response.write "selected" %>>�������ݼ�</option>
		<option value="best" <% IF sortDiv="best" Then response.write "selected" %>>����Ʈ��</option>
		</select>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td align="center">����</td>
	<td align="center">��ǰID</td>
	<td align="center">�̹���</td>
	<td align="center">��ǰ��</td>
	<td align="center">����</td>
	<td align="center">��üID</td>
	<td align="center">��۱���</td>
	<td align="center">�Ǹſ���</td>
</tr>
<% for i=0 to oItem.FresultCount-1 %>
<form name="frmBuyPrc_<%= oItem.FItemList(i).FItemID %>" method="post" onSubmit="return false;" action="doitemviewset.asp">
<input type="hidden" name="mode" value="">
<input type="hidden" name="itemid" value="<%= oItem.FItemList(i).FItemID %>">
<input type="hidden" name="itemname" value="<%= oItem.FItemList(i).FItemName %>">
<tr bgcolor="#FFFFFF" align="center">
	<td><input type="radio" name="cksel" onClick="AnCheckClick(this);"></td>
	<td><%= oItem.FItemList(i).FItemID %></td>
	<td><img src="<%= oItem.FItemList(i).Fsmallimage %>" width="50" height="50" border=0 alt=""></td>
	<td align="left"><%= oItem.FItemList(i).FItemName %></td>
	<td align="left"><%= FormatNumber(oItem.FItemList(i).Fsellcash,0) %></td>
	<td align="left"><%= oItem.FItemList(i).FMakerID %></td>
	<td>
	<% if oItem.FItemList(i).IsUpcheBeasong() then Response.Write "��ü���":Else Response.Write "10X10" %>
	</td>
	<td>
	<% if oItem.FItemList(i).FSellYn="Y" then %>
	Y
	<% else %>
	N
	<% end if %>
	</td>
</tr>
</form>
<% next %>
<tr bgcolor="#FFFFFF">
	<td colspan="9" align="center">
	<% if oItem.HasPreScroll then %>
		<a href="?page=<%= oItem.StartScrollPage-1 %>&itemid=<%= itemid %>&itemname=<%= itemname %>&makerid=<%= makerid %>&sellyn=<%= sellyn %>&packyn=<%=packyn%>&target=<%= target %>&cdl=<% = cdl %>&cdm=<% = cdm %>&cds=<% = cds %>&sd=<%=sortdiv%>">[pre]</a>
	<% else %>
		[pre]
	<% end if %>

	<% for i=0 + oItem.StartScrollPage to oItem.FScrollCount + oItem.StartScrollPage - 1 %>
		<% if i>oItem.FTotalpage then Exit for %>
		<% if CStr(page)=CStr(i) then %>
		<font color="red">[<%= i %>]</font>
		<% else %>
		<a href="?page=<%= i %>&itemid=<%= itemid %>&itemname=<%= itemname %>&makerid=<%= makerid %>&sellyn=<%= sellyn %>&packyn=<%=packyn%>&target=<%= target %>&cdl=<% = cdl %>&cdm=<% = cdm %>&cds=<% = cds %>&sd=<%=sortdiv%>">[<%= i %>]</a>
		<% end if %>
	<% next %>

	<% if oItem.HasNextScroll then %>
		<a href="?page=<%= i %>&itemid=<%= itemid %>&itemname=<%= itemname %>&makerid=<%= makerid %>&sellyn=<%= sellyn %>&packyn=<%=packyn%>&target=<%= target %>&cdl=<% = cdl %>&cdm=<% = cdm %>&cds=<% = cds %>&sd=<%=sortdiv%>">[next]</a>
	<% else %>
		[next]
	<% end if %>
	</td>
</tr>
</table>
<%
set oItem = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
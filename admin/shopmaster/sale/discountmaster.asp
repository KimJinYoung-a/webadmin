<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/items/discountitemcls.asp"-->
<%
'####################################################
' Description :  ���� ��ǰ ����
' History : 2008.04.07 ������ ����
'####################################################

dim sCode,eCode
dim sBrand,itemid,itemidArr
dim malltype, page

sCode     	= requestCheckVar(Request("sC"),10)
eCode		= requestCheckVar(Request("eC"),10)
sBrand		= requestCheckVar(request("ebrand"),32)

malltype = request("malltype")
itemid = request("itemid")
page = request("page")
itemidArr = Trim(request("itemidArr"))

if Right(itemidArr,1)="," then itemidArr=Left(itemidArr,Len(itemidArr)-1)

if page="" then page="1"

dim odiscount
set odiscount = new CDiscount
odiscount.FPageSize=30
odiscount.FCurrPage= page
odiscount.FRectMallType = malltype
odiscount.FRectItemID = itemid
odiscount.FRectitemidArr = itemidArr
odiscount.FRectDesingerID = sBrand

if (sBrand<>"") or (itemid<>"") or (itemidArr<>"") then
	odiscount.GetDesignerItemList
end if

dim i

%>
<script language='javascript'>
function SelectCk(opt){
	var bool = opt.checked;
	AnSelectAllFrame(bool)
}


function orgprice(iitemid){
	var ret = confirm('������ �����Ͻðڽ��ϱ�?');

	var frm = document.frmorg;
	if (ret){
		frm.itemid.value = iitemid;
		frm.submit();
	}
}

function CkDisOrOrg(isDisc){
	var frm;
	var pass = false;

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			pass = ((pass)||(frm.cksel.checked));
		}
	}

	var ret;

	if (!pass) {
		alert('���� �������� �����ϴ�.');
		return;
	}


	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			if (frm.cksel.checked){
				if (isDisc==true){
					frm.sailyn[0].checked=true;
				}else{
					frm.sailyn[1].checked=true;
				}
			}
		}
	}
}

function CkDisPrice(){
	CkDisOrOrg(true);
}

function CkOrgPrice(){
	CkDisOrOrg(false);
}

function sailProAct(){
	var frm;
	var pass = false;
	var sailpro = document.frmdummi.sailpro.value;

	if (!IsDigit(sailpro)){
		alert('���ڸ� �����մϴ�.');
		document.frmdummi.sailpro.focus();
		return;
	}

	if (sailpro*1>99){
		alert('100���� ���ڸ� �����մϴ�.');
		document.frmdummi.sailpro.focus();
		return;
	}

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			pass = ((pass)||(frm.cksel.checked));
		}
	}

	var ret;

	if (!pass) {
		alert('���� �������� �����ϴ�.');
		return;
	}


	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			if (frm.cksel.checked){
				frm.dissellprice.value = parseInt(Math.round(frm.orgprice.value * (1 - sailpro/100.0)));
			}
		}
	}
}

function sailMargineAct(){
	var frm;
	var pass = false;
	var maeippro = document.frmdummi.maeippro.value;

	if (!IsDigit(maeippro)){
		alert('���ڸ� �����մϴ�.');
		document.frmdummi.maeippro.focus();
		return;
	}

	if (maeippro*1>99){
		alert('100���� ���ڸ� �����մϴ�.');
		document.frmdummi.maeippro.focus();
		return;
	}

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			pass = ((pass)||(frm.cksel.checked));
		}
	}

	var ret;

	if (!pass) {
		alert('���� �������� �����ϴ�.');
		return;
	}


	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			if (frm.cksel.checked){
				frm.disbuyprice.value = parseInt(Math.round(frm.dissellprice.value * (1 - maeippro/100.0)));
			}
		}
	}
}

function saveArr(){
	var frm;
	var pass = false;

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			pass = ((pass)||(frm.cksel.checked));
		}
	}

	var ret;

	if (!pass) {
		alert('���� �������� �����ϴ�.');
		return;
	}

	frmarr.itemid.value = "";
	frmarr.sailyn.value = "";
	frmarr.dissellprice.value = "";
	frmarr.disbuyprice.value = "";

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			if (frm.cksel.checked){
				//check Not AvaliValue
				if (!IsDigit(frm.dissellprice.value)){
					alert('���ڸ� �����մϴ�.');
					frm.dissellprice.focus();
					return;
				}

				if (frm.dissellprice.value<1){
					alert('�ݾ��� ��Ȯ�� �Է��ϼ���.');
					frm.dissellprice.focus();
					return;
				}

				if (!IsDigit(frm.disbuyprice.value)){
					alert('���ڸ� �����մϴ�.');
					frm.disbuyprice.focus();
					return;
				}

				if (frm.disbuyprice.value<1){
					alert('�ݾ��� ��Ȯ�� �Է��ϼ���.');
					frm.disbuyprice.focus();
					return;
				}

				frmarr.itemid.value = frmarr.itemid.value + frm.itemid.value + ","
				if (frm.sailyn[0].checked){
					frmarr.sailyn.value = frmarr.sailyn.value + "Y" + ","
				}else{
					frmarr.sailyn.value = frmarr.sailyn.value + "N" + ","
				}
				frmarr.dissellprice.value = frmarr.dissellprice.value + frm.dissellprice.value + ","
				frmarr.disbuyprice.value = frmarr.disbuyprice.value + frm.disbuyprice.value + ","


			}
		}
	}

	var ret = confirm('���� �Ͻðڽ��ϱ�?');

	if (ret){
		frmarr.submit();
	}

}
</script>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frmSearch" method="get" action="">
	<input type="hidden" name="page" value="1">
	<input type="hidden" name="menupos" value="<%=menupos%>">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
		&nbsp;�귣��:
		<% drawSelectBoxDesignerwithName "ebrand", sBrand%>
		itemid :
		<input type="text" name="itemid" value="<%= itemid %>" size="9" maxlength="9">
        <br>
        itemid �޸� ���� :
        <input type="text" name="itemidArr" value="<%= itemidArr %>" size="100" maxlength="200">
	<td  width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="�˻�" onClick="javascript:document.frmSearch.submit();">
	</td>
	</tr>
	</form>
</table>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" border=0>
<form name=frmdummi>
<tr height="40" valign="bottom">
	<td align="left"><input type=button value="���û�ǰ����" onClick="saveArr()" class="button"></td>
	<td align="right">
	<input type="button" value="�����Ǹ�" onClick="CkDisPrice();" class="button">
	&nbsp;&nbsp;
	<input type="button" value="�����Ǹ�" onClick="CkOrgPrice();" class="button">
	&nbsp;&nbsp;
	���ǸŰ��� <input type=text name=sailpro value="" size=2 maxlength=2>% ����
	<input type=button value="����" onclick="sailProAct()" class="button">&nbsp;&nbsp;
	���θ����� <input type=text name=maeippro value="" size=2 maxlength=2>%<input type=button value="����" onclick="sailMargineAct()" class="button">
	</td>
</tr>
</form>
</table>
<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="<%= adminColor("tabletop") %>">
	<td align="center" width="20"><input type="checkbox" name="ck_all" onclick="SelectCk(this)"></td>
	<td align="center" width="40">��ǰID</td>
	<td align="center" width="50" >�̹���</td>
	<td align="center">��ǰ��</td>
	<td align="center" width="80">�귣��</td>
	<td align="center" width="40">���<br>����</td>
	<td align="center" width="40">����</td>
	<td align="center" width="50">����<br>�ǸŰ�</td>
	<td align="center" width="50">����<br>���԰�</td>
	<td align="center" width="50">����<br>������</td>

	<td align="center" width="50">��<br>�ǸŰ�</td>
	<td align="center" width="50">��<br>���԰�</td>
	<td align="center" width="50">��<br>������</td>

	<td align="center" width="50">����<br>�ǸŰ�</td>
	<td align="center" width="50">����<br>���԰�</td>
	<td align="center" width="50">����<br>������</td>

	<!-- <td align="center">����</td> -->
</tr>
<% for i=0 to odiscount.FResultCount -1 %>
<form name="frmBuyPrc_<%= odiscount.FItemList(i).FItemID %>" >
<input type=hidden name=orgprice value="<%= odiscount.FItemList(i).Forgprice %>">
<input type=hidden name=itemid value="<%= odiscount.FItemList(i).FItemID %>">
<% if odiscount.FItemList(i).FSailYn="Y" then %>
<tr bgcolor="#CCCCCC">
<% else %>
<tr bgcolor="#FFFFFF">
<% end if %>
	<td><input type="checkbox" name="cksel" onClick="AnCheckClick(this);" ></td>
	<td align="center"><%= odiscount.FItemList(i).FItemID %></td>
	<td><img src="<%= odiscount.FItemList(i).FImageSmall %>" height="50" width="50"></td>
	<td><%= odiscount.FItemList(i).FItemName %></td>
	<td align="center"><%= odiscount.FItemList(i).FMakerID %></td>
	<td align="center">
	<%
		Select Case odiscount.FItemList(i).Fmwdiv
			Case "M"
				Response.Write "<Font color=#F08050>����</font>"
			Case "W"
				Response.Write "<Font color=#808080>��Ź</font>"
			Case "U"
				Response.Write "<Font color=#5080F0>��ü</font>"
		end Select
	%>
	</td>
	<td align="center">
	<% if odiscount.FItemList(i).FSailYn="Y" then %>
	<input type=radio name=sailyn value="Y" checked ><font color=red>Y</font>
	<input type=radio name=sailyn value="N" >N
	<% else %>
	<input type=radio name=sailyn value="Y" >Y
	<input type=radio name=sailyn value="N" checked >N
	<% end if %>
	</td>
	<td align="right"><%= FormatNumber(odiscount.FItemList(i).FSellcash,0) %></td>
	<td align="right"><%= FormatNumber(odiscount.FItemList(i).FBuycash,0) %></td>
	<td align="center">
	<% if odiscount.FItemList(i).FSellcash<>0 then %>
	<%= 100-fix(odiscount.FItemList(i).FBuycash/odiscount.FItemList(i).FSellcash*10000)/100 %>%
	<% end if %>
	</td>
	<td align="right"><%= FormatNumber(odiscount.FItemList(i).Forgprice,0) %></td>
	<td align="right"><%= FormatNumber(odiscount.FItemList(i).Forgsuplycash,0) %></td>
	<td align="center">
	<% if odiscount.FItemList(i).Forgprice<>0 then %>
	<%= 100-fix(odiscount.FItemList(i).Forgsuplycash/odiscount.FItemList(i).Forgprice*10000)/100 %>%
	<% end if %>
	</td>
	<td align="right">
	<input type=text name=dissellprice value="<%= odiscount.FItemList(i).Fsailprice %>" size=6 maxlength=9>
	</td>
	<td align="right">
	<input type=text name=disbuyprice value="<%= odiscount.FItemList(i).Fsailsuplycash %>" size=6 maxlength=9>
	</td>
	<td align="center">
	<% if odiscount.FItemList(i).Fsailprice<>0 then %>
	<%= 100-fix(odiscount.FItemList(i).Fsailsuplycash/odiscount.FItemList(i).Fsailprice*10000)/100 %>%
	<% end if %>
	</td>
<!--
	<td><input type="button" value="������" onClick="orgprice('<%= odiscount.FItemList(i).FItemID %>')"></td>
-->
</tr>
</form>
<% next %>
<tr bgcolor="#FFFFFF">
	<td colspan="16" align="center">
	<% if odiscount.HasPreScroll then %>
		<a href="?page=<%= odiscount.StarScrollPage-1 %>&itemid=<%= itemid %>&malltype=<%= malltype %>&sBrand=<%= sBrand %>">[pre]</a>
	<% else %>
		[pre]
	<% end if %>

	<% for i=0 + odiscount.StarScrollPage to odiscount.FScrollCount + odiscount.StarScrollPage - 1 %>
		<% if i>odiscount.FTotalpage then Exit for %>
		<% if CStr(page)=CStr(i) then %>
		<font color="red">[<%= i %>]</font>
		<% else %>
		<a href="?page=<%= i %>&itemid=<%= itemid %>&malltype=<%= malltype %>&sBrand=<%= sBrand %>">[<%= i %>]</a>
		<% end if %>
	<% next %>

	<% if odiscount.HasNextScroll then %>
		<a href="?page=<%= i %>&itemid=<%= itemid %>&malltype=<%= malltype %>&sBrand=<%= sBrand %>">[next]</a>
	<% else %>
		[next]
	<% end if %>
	</td>
</tr>
</table>
<form name=frmarr method=post action="dodiscountitem.asp">
<input type=hidden name=mode value="arrdischange">
<input type=hidden name=itemid value="">
<input type=hidden name=sailyn value="">
<input type=hidden name=dissellprice value="">
<input type=hidden name=disbuyprice value="">
</form>
<%
set odiscount = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
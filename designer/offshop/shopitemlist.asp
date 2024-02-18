<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/designer/lib/designerbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshopitemcls.asp"-->
<%
dim designer,page,ckonlyoff,ckonlyusing,research
designer    = reQuestCheckVar(session("ssBctID"),32)
page        = reQuestCheckVar(request("page"),10)
ckonlyoff   = reQuestCheckVar(request("ckonlyoff"),10)
ckonlyusing = reQuestCheckVar(request("ckonlyusing"),10)
research    = reQuestCheckVar(request("research"),10)

if page="" then page=1
if research<>"on" then ckonlyusing="Y"

dim ioffitem
set ioffitem  = new COffShopItem
ioffitem.FPageSize = 50
ioffitem.FCurrPage = page
ioffitem.FRectDesigner = designer
ioffitem.FRectOnlyOffLine = ckonlyoff
ioffitem.FRectOnlyUsing = ckonlyusing

if designer<>"" then
	ioffitem.GetOffShopItemList
end if

dim i

dim IsDirectIpchulContractExistsBrand
IsDirectIpchulContractExistsBrand = fnIsDirectIpchulContractExistsBrand(designer)

%>
<script language='javascript'>
function NextPage(page){
	frm.page.value=page;
	frm.submit();
}
function popOffItemEdit(ibarcode){
    <% if Not (IsDirectIpchulContractExistsBrand) then %>
        alert('������ �����ϴ�. - ���� ���� �԰� �귣�常 ���� �����մϴ�.');
        return;
    <% end if %>
	var popwin = window.open('popoffitemedit.asp?barcode=' + ibarcode,'offitemedit','width=500,height=800,resizable=yes,scrollbars=yes');
	popwin.focus();
}


function OffItemReg(idesigner){
    <% if Not (IsDirectIpchulContractExistsBrand) then %>
        alert('������ �����ϴ�. - ���� ���� �԰� �귣�常 ���� �����մϴ�.');
        return;
    <% end if %>
	var subwin;
	subwin = window.open('popoffitemreg.asp?designer=' + idesigner,'window_reg','width=500,height=600,scrollbars=yes,resizable=yes');
	subwin.focus();
}

function AnSearch(frm){
	frm.submit();
}

function CheckThis(frm){
	frm.cksel.checked=true;
	AnCheckClick(frm.cksel);
}

function SelectCk(opt){
	var bool = opt.checked;
	AnSelectAllFrame(bool)
}

function ChargeIdAvail(ichargeid){
	var comp = document.frm.designer;

	if (ichargeid=="10x10"){
		return true
	}

	for (var i=0;i<comp.length;i++){
		if (comp[i].value==ichargeid){
			return true
		}
	}

	return false;
}

function ModiArr(){
    <% if Not (IsDirectIpchulContractExistsBrand) then %>
        alert('������ �����ϴ�. - ���� ���� �԰� �귣�常 ���� �����մϴ�.');
        return;
    <% end if %>

	var upfrm = document.frmArrupdate;
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

	upfrm.itemgubunarr.value = "";
	upfrm.itemarr.value = "";
	upfrm.itemoptionarr.value = "";
	upfrm.isusingarr.value = "";
	upfrm.extbarcodearr.value = "";

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			if (frm.cksel.checked){
				upfrm.itemgubunarr.value = upfrm.itemgubunarr.value + frm.itemgubun.value + "|";
				upfrm.itemarr.value = upfrm.itemarr.value + frm.itemid.value + "|";
				upfrm.itemoptionarr.value = upfrm.itemoptionarr.value + frm.itemoption.value + "|";
				upfrm.extbarcodearr.value = upfrm.extbarcodearr.value + frm.extbarcode.value + "|";

				if (frm.isusing[0].checked){
					upfrm.isusingarr.value = upfrm.isusingarr.value + "Y" + "|";
				}else{
					upfrm.isusingarr.value = upfrm.isusingarr.value + "N" + "|";
				}
			}
		}
	}

	var ret = confirm('���� �Ͻðڽ��ϱ�?');

	if (ret){
		upfrm.mode.value = "arrmodi";
		upfrm.submit();
	}
}

</script>

<!-- ���Ŀ� �޴�����κп� �־�� �մϴ�. -->
<table width="100%" border="0" valign="top" cellpadding="0" cellspacing="0" class="a">
	<tr bgcolor="#FFFFFF">
		<td style="padding:5; border:1px solid <%= adminColor("tablebg") %>;" bgcolor="#FFFFFF">
			* ������ �����ǰ�� ���� �̹��� ����� �ʼ��� ����Ǿ����ϴ�.<br>
			* ��Ȱ�� �ֹ� ����ó���� ���� �̹��� ���� ��ǰ�� ���� <b>�̹����� ���</b>�� �ּ���<br>
			* ��ǰ�������� �����Ϸ��� ��ǰ��ȣ�� �����ּ���.
		</td>
	</tr>
</table>
<!-- ���Ŀ� �޴�����κп� �־�� �մϴ�. -->

<p>


<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="page" value="">
	<input type="hidden" name="research" value="on">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
		<td align="left">
			���:<% drawSelectBoxUsingYN "ckonlyusing", ckonlyusing %>
			&nbsp;
			��ǰ����:
			<select class="select" name="ckonlyoff">
		     	<option value='' selected>��ü</option>
		     	<option value='10' <% if ckonlyoff="10" then response.write "selected" %>>�¶��ε�ϻ�ǰ(10)</option>
		     	<option value='90' <% if ckonlyoff="90" then response.write "selected" %>>���������ǰ(90)</option>
	     	</select>
		</td>

		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="�˻�" onClick="javascript:document.frm.submit();">
		</td>
	</tr>
	</form>
</table>


<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<tr>
		<td align="left">
			<input type="button" class="button" value="������������ ��ǰ���" onClick="OffItemReg('<%=designer%>')">
			<% if ioffitem.FresultCount>0 then %>
			<input type="button" class="button" value="���þ���������" onclick="ModiArr()">
			<% end if %>

			<% if Not (IsDirectIpchulContractExistsBrand) then %>
            (���� ���� �԰� �귣�常 ���� �����մϴ�.)
            <% end if %>
		</td>
		<td align="right">

		</td>
	</tr>
</table>
<!-- �׼� �� -->

<p>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15">
			�˻���� : <b><%= ioffitem.FTotalCount%></b>
			&nbsp;
			������ : <b><%= page %> / <%=  ioffitem.FTotalpage %></b>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    	<td width="20"><input type="checkbox" name="ckall" onClick="SelectCk(this)"></td>
    	<td width="50">�̹���</td>
    	<td width="100">��ǰ�ڵ�</td>
    	<td>��ǰ��</td>
    	<td>�ɼǸ�</td>
    	<td width="60">�Һ��ڰ�</td>
    	<td width="60">�ǸŰ�</td>
    	<td width="110">������ڵ�</td>
    	<td width="80">��뿩��</td>
	</tr>
	<% for i=0 to ioffitem.FresultCount -1 %>
	<form name="frmBuyPrc_<%= i %>" >
	<input type="hidden" name="itemgubun" value="<%= ioffitem.FItemlist(i).Fitemgubun %>">
	<input type="hidden" name="itemid" value="<%= ioffitem.FItemlist(i).Fshopitemid %>">
	<input type="hidden" name="itemoption" value="<%= ioffitem.FItemlist(i).Fitemoption %>">
	<input type="hidden" name="tx_charge" value="<%= ioffitem.FItemlist(i).FMakerID %>">
	<tr align="center" bgcolor="#FFFFFF">
  		<td ><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"></td>
  		<td ><a href="javascript:popOffItemEdit('<%= ioffitem.FItemlist(i).GetBarCode %>');"><img src="<%= ioffitem.FItemlist(i).GetImageSmall %>" width=50 height=50 onError="this.src='http://webimage.10x10.co.kr/images/no_image.gif'" border=0></a></td>
  		<td><a href="javascript:popOffItemEdit('<%= ioffitem.FItemlist(i).GetBarCode %>');"><%= ioffitem.FItemlist(i).Fitemgubun %>-<%= CHKIIF(ioffitem.FItemlist(i).Fshopitemid>=1000000,Format00(8,ioffitem.FItemlist(i).Fshopitemid),Format00(6,ioffitem.FItemlist(i).Fshopitemid)) %>-<%= ioffitem.FItemlist(i).Fitemoption %></a></td>
  		<td align="left"><%= ioffitem.FItemlist(i).FShopItemName %></td>
  		<td><%= ioffitem.FItemlist(i).FShopitemOptionname %></td>
  		<td align="right" ><%= FormatNumber(ioffitem.FItemlist(i).FShopItemOrgprice,0) %></td>
  		<td align="right" ><%= FormatNumber(ioffitem.FItemlist(i).FShopItemprice,0) %></td>
  		<td><input type="text" name="extbarcode" value="<%= ioffitem.FItemlist(i).FextBarcode %>" size="13" maxlength="32" style="border:1px #999999 solid;" onKeyPress="CheckThis(frmBuyPrc_<%= i %>)"></td>
  		<td align="center" >
  		<% if ioffitem.FItemlist(i).Fisusing="Y" then %>
  		<input type="radio" name="isusing" value="Y" checked onclick="CheckThis(frmBuyPrc_<%= i %>)">Y
  		<input type="radio" name="isusing" value="N" onclick="CheckThis(frmBuyPrc_<%= i %>)">N
  		<% else %>
  		<input type="radio" name="isusing" value="Y" onclick="CheckThis(frmBuyPrc_<%= i %>)">Y
  		<input type="radio" name="isusing" value="N" checked onclick="CheckThis(frmBuyPrc_<%= i %>)"><font color="red">N</font>
  		<% end if %>
  		</td>
  	</tr>
  	</form>
  	<% next %>
  	<tr bgcolor="#FFFFFF">
		<td colspan="10" align="center">
		<% if ioffitem.HasPreScroll then %>
			<a href="javascript:NextPage('<%= ioffitem.StartScrollPage-1 %>');">[pre]</a>
		<% else %>
			[pre]
		<% end if %>

		<% for i=0 + ioffitem.StartScrollPage to ioffitem.FScrollCount + ioffitem.StartScrollPage - 1 %>
			<% if i>ioffitem.FTotalpage then Exit for %>
			<% if CStr(page)=CStr(i) then %>
			<font color="red">[<%= i %>]</font>
			<% else %>
			<a href="javascript:NextPage('<%= i %>');">[<%= i %>]</a>
			<% end if %>
		<% next %>

		<% if ioffitem.HasNextScroll then %>
			<a href="javascript:NextPage('<%= i %>')">[next]</a>
		<% else %>
			[next]
		<% end if %>
		</td>
	</tr>
</table>
<form name="frmArrupdate" method="post" action="doshopitem.asp">
<input type="hidden" name="mode" value="">
<input type="hidden" name="itemgubunarr" value="">
<input type="hidden" name="itemarr" value="">
<input type="hidden" name="itemoptionarr" value="">
<input type="hidden" name="isusingarr" value="">
<input type="hidden" name="extbarcodearr" value="">
</form>
<%
set ioffitem  = Nothing
%>
<!-- #include virtual="/designer/lib/designerbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
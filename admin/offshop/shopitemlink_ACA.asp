<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  �������� ��ǰ ���
' History : 2009.04.07 ������ ����
'			2010.12.13 �ѿ�� ����
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/commonbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshopitemcls.asp"-->
<%
dim designer,page ,cdl, cdm, cds ,itemid , i ,IsDirectIpchulContractExistsBrand, itemname
dim research, mwdiv, sellyn, usingyn, itemlinktype
	designer    = RequestCheckVar(request("designer"),32)
	page        = RequestCheckVar(request("page"),9)
	research    = RequestCheckVar(request("research"),9)
	mwdiv       = RequestCheckVar(request("mwdiv"),9)
	usingyn     = RequestCheckVar(request("usingyn"),9)
	sellyn      = RequestCheckVar(request("sellyn"),9)
	cdl         = RequestCheckVar(request("cdl"),3)
	cdm         = RequestCheckVar(request("cdm"),3)
	cds         = RequestCheckVar(request("cds"),3)
	itemid      = RequestCheckVar(request("itemid"),9)
	itemname    = RequestCheckVar(request("itemname"),32)

if page="" then page=1

if (research="") and (mwdiv="") then mwdiv="MW"  ''�⺻��. MW (��ü�� �ƴѰ��)
if (research="") and (usingyn="") then usingyn="Y"  ''�⺻��. Y
if (research="") and (sellyn="") then sellyn="Y"  ''�⺻��. Y

if (itemid <> "") then
	if Not IsNumeric(itemid) then
		response.write "<script>alert('�߸��� ��ǰ�ڵ��Դϴ�. : " & itemid & "');</script>"
		itemid = ""
	else
		itemid = CLng(itemid)
	end if
end if


dim ioffitem
set ioffitem  = new COffShopItem
	ioffitem.FPageSize = 100
	ioffitem.FCurrPage = page
	ioffitem.FRectDesigner = designer
	ioffitem.FRectitemid = itemid
	ioffitem.FRectOnlineMWdiv = mwdiv
	ioffitem.FRectIsusing = usingyn
	ioffitem.FRectSellYN  = sellyn
	''ioffitem.FRectitemname  = itemname
	''ioffitem.FRectcdl  = cdl
	''ioffitem.FRectcdm  = cdm
	''ioffitem.FRectcds  = cds
	ioffitem.GetAcaLinkReqList()

	''response.write ioffitem.FTotalCount

'if itemlinktype	= "" then itemlinktype = "O"
%>
<script type="text/javascript">

function frmsubmit() {
	/*
	if (frm.itemname.value!='') {
		if (frm.designer.value=='') {
			alert('��ǰ�� �˻��� �귣�带 �ݵ�� �־� �ּ���.(���Ϲ���)');
			return;
		}
	}
	*/

	if(frm.itemid.value!='') {
		if (!IsDouble(frm.itemid.value)) {
			alert('��ǰ��ȣ�� ���ڸ� �����մϴ�.');
			frm.itemid.focus();
			return;
		}
	}

	frm.submit();
}

function SelectCk(opt) {
	var bool = opt.checked;
	AnSelectAllFrame(bool)
}


function AddArr() {
	var upfrm = document.frmArrupdate;
	var frm;
	var pass = false;

	for (var i=0;i<document.forms.length;i++) {
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
	upfrm.itemidarr.value = "";
	upfrm.itemoptionarr.value = "";

	for (var i=0;i<document.forms.length;i++) {
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			if (frm.cksel.checked) {
				upfrm.itemgubunarr.value = upfrm.itemgubunarr.value + frm.itemgubun.value + "|";
				upfrm.itemidarr.value = upfrm.itemidarr.value + frm.itemid.value + "|";
				upfrm.itemoptionarr.value = upfrm.itemoptionarr.value + frm.itemoption.value + "|";
			}
		}
	}
	var ret = confirm('���� �Ͻðڽ��ϱ�?');

	if (ret) {
		upfrm.mode.value = "arraddACA";
		upfrm.submit();
	}
}

function CheckThis(frm) {
	frm.cksel.checked=true;
	AnCheckClick(frm.cksel);
}

function gotoPage(page) {
	document.frm.page.value = page;
	document.frm.submit();
}

</script>

<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="research" value="on">
<input type="hidden" name="page" value="1">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
		��ǰ�ڵ� : <input type="text" class="text" name="itemid" value="<%= itemid %>" size="7" maxlength="9">
		&nbsp;
		�귣�� :
		<% drawSelectBoxDesignerwithName "designer",designer  %>
	</td>
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="frmsubmit();">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
	    <span style="white-space:nowrap;">ON ���Ա���:<% drawSelectBoxMWU "mwdiv", mwdiv %></span>
	    &nbsp;
	    <span style="white-space:nowrap;">ON �Ǹſ���:<% drawSelectBoxSellYN "sellyn", sellyn %></span>
	    &nbsp;
        <span style="white-space:nowrap;">ON ��뿩��:<% drawSelectBoxUsingYN "usingyn", usingyn %></span>

	    <% if (FALSE) then %>
		<input type="radio" name="umwdiv" value="ALL" <% if umwdiv="ALL" then response.write "checked" %> <% if designer = "" then response.write " disabled" %>>(�Ǹ���)����ǰ
		<input type="radio" name="umwdiv" value="Y" <% if umwdiv="Y" then response.write "checked" %> <% if designer = "" then response.write " disabled" %>>(�Ǹ���)��ü��ۻ�ǰ
		<input type="radio" name="umwdiv" value="N" <% if umwdiv="N" then response.write "checked" %> <% if designer = "" then response.write " disabled" %>>�Ǹ����� ��ǰ �˻�
	    <% end if %>
	</td>
</tr>
</form>
</table>
<!-- �˻� �� -->

<br>

<!-- �׼� ���� -->
�� �ΰŽ� ��ǰ �� �������� ��ǰ���� ��ϵ��� ���� ��ǰ ����Ʈ �Դϴ�.<br>
�� ����Ͻø� [������ǰ����] �޴��� ��ǰ�� ��Ÿ���� ���ڵ� ��� �Ͻ� �� �ֽ��ϴ�.
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td align="left">
	</td>
	<td align="right">
		<% if ioffitem.FresultCount>0 then %>
		<input type="button" class="button" value="���� ��ǰ �������� ��ǰ���� ���" onclick="AddArr()">
		<% end if %>
	</td>
</tr>
</table>
<!-- �׼� �� -->

<!-- ����Ʈ ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		�˻���� : <b><%= ioffitem.FTotalCount %></b>
		&nbsp;
		������ : <b><%= Page %> / <%= ioffitem.FTotalPage %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="20"><input type="checkbox" name="ckall" onClick="SelectCk(this)"></td>
	<td width="100">��ǰ�ڵ�</td>
	<td>�귣��ID</td>
	<td>��ǰ��<font color="blue">[�ɼǸ�]</font></td>
	<td width="50">�ΰŽ�<br>���Ա���</td>
	<!-- <td width="50">Center<br>���Ա���</td> -->
	<td width="90">�Һ��ڰ�</td>
	<td width="90">�ǸŰ�</td>
	<td width="250">���</td>
</tr>

<% if ioffitem.FresultCount > 0 then %>
	<%
	for i=0 to ioffitem.FresultCount -1

	''�������̸鼭 �Ⱓ������ �ƴϸ� �ǸŰ��� ���(�������)
	if ioffitem.FItemlist(i).Ftermsale ="N" and ioffitem.FItemlist(i).FOnlineitemorgprice>ioffitem.FItemlist(i).FShopItemprice then
		itemlinktype = "S"
	else
		itemlinktype = "O"
	end if
	%>
	<form name="frmBuyPrc_<%= i %>" >
	<input type="hidden" name="itemgubun" value="<%= ioffitem.FItemlist(i).Fitemgubun %>">
	<input type="hidden" name="itemid" value="<%= ioffitem.FItemlist(i).Fshopitemid %>">
	<input type="hidden" name="itemoption" value="<%= ioffitem.FItemlist(i).Fitemoption %>">
	<input type="hidden" name="makerid" value="<%= ioffitem.FItemlist(i).FMakerID %>">
	<tr align="center" bgcolor="#FFFFFF">
		<td><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"></td>
		<td><%= ioffitem.FItemlist(i).Fitemgubun %>-<%= CHKIIF(ioffitem.FItemlist(i).Fshopitemid>=1000000,Format00(8,ioffitem.FItemlist(i).Fshopitemid),Format00(6,ioffitem.FItemlist(i).Fshopitemid)) %>-<%= ioffitem.FItemlist(i).Fitemoption %></td>
		<td>
			<%= ioffitem.FItemlist(i).FMakerID %>
		</td>
		<td align="left">
			<%= ioffitem.FItemlist(i).FShopItemName %>
			<% if ioffitem.FItemlist(i).Fitemoption<>"0000" then %>
				<font color="blue">[<%= ioffitem.FItemlist(i).FShopitemOptionname %>]</font>
			<% end if %>
		</td>
		<td><font color="<%= ioffitem.FItemlist(i).getMwDivColor %>"><%= ioffitem.FItemlist(i).getMwDivName %></font></td>
		<!-- <td></td> -->
		<td align="right"><%= FormatNumber(ioffitem.FItemlist(i).FOnlineitemorgprice,0) %></td>
		<td align="right">
			<% if ioffitem.FItemlist(i).Ftermsale ="Y" then %>
				<font color="red">�Ⱓ����</font>
		    <% elseif (ioffitem.FItemlist(i).FOnlineitemorgprice>ioffitem.FItemlist(i).FShopItemprice) then %>
		    	<font color="red"><!--���-->����</font>
		    <% end if %>
		    <%= FormatNumber(ioffitem.FItemlist(i).FShopItemprice,0) %>
		</td>
		<td>
		</td>
	</tr>
	</form>
	<% next %>

    <tr align="center" bgcolor="#FFFFFF">
      	<td colspan="8">
	   	<% if ioffitem.HasPreScroll then %>
			<span class="list_link"><a href="javascript:gotoPage(<%= ioffitem.StartScrollPage-1 %>)">[pre]</a></span>
		<% else %>
		[pre]
		<% end if %>
		<% for i = 0 + ioffitem.StartScrollPage to ioffitem.StartScrollPage + ioffitem.FScrollCount - 1 %>
			<% if (i > ioffitem.FTotalpage) then Exit for %>
			<% if CStr(i) = CStr(ioffitem.FCurrPage) then %>
			<span class="page_link"><font color="red"><b>[<%= i %>]</b></font></span>
			<% else %>
			<a href="javascript:gotoPage(<%= i %>)" class="list_link"><font color="#000000">[<%= i %>]</font></a>
			<% end if %>
		<% next %>
		<% if ioffitem.HasNextScroll then %>
			<span class="list_link"><a href="javascript:gotoPage(<%= i %>)">[next]</a></span>
		<% else %>
		[next]
		<% end if %>
      	</td>
    </tr>
<% else %>
	<tr align="center" bgcolor="#FFFFFF">
		<td colspan=20>�˻������ �����ϴ�</td>
	</tr>
<% end if %>

<form name="frmArrupdate" method="post" action="shopitem_process.asp">
	<input type="hidden" name="mode" value="">
	<input type="hidden" name="itemgubunarr" value="">
	<input type="hidden" name="itemidarr" value="">
	<input type="hidden" name="itemoptionarr" value="">
</form>
</table>
<%
set ioffitem  = Nothing
%>
<!-- #include virtual="/common/lib/commonbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->

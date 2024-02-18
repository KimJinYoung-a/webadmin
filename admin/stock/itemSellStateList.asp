<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  �귣�庰�����Ȳ
' History : 2009.04.07 ������ ����
'			2013.10.16 �ѿ�� ����
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stock/summary_itemstockcls.asp"-->
<!-- #include virtual="/lib/BarcodeFunction.asp"-->
<%

dim makerid, itemgubun, itemidArr, itemoption
dim sellyn, isusing, optsellyn, optisosing, limitrealstock, stocktype, mwdiv
dim pagesize, page, ordby, research
dim i, j, k


research        = requestCheckvar(request("research"),9)
makerid			= requestCheckvar(request("makerid"),32)
itemgubun       = requestCheckvar(request("itemgubun"),32)
itemidArr       = requestCheckvar(request("itemidArr"),3200)
sellyn         	= requestCheckvar(request("sellyn"),32)
isusing         = requestCheckvar(request("isusing"),32)
optsellyn       = requestCheckvar(request("optsellyn"),32)
optisosing      = requestCheckvar(request("optisosing"),32)
limitrealstock  = requestCheckvar(request("limitrealstock"),32)
mwdiv    		= requestCheckvar(request("mwdiv"),64)
ordby    		= requestCheckvar(request("ordby"),64)



if (research = "") then sellyn = "NS"
if (pagesize = "") then pagesize = 200
if (page = "") then page = 1
if (limitrealstock = "") then limitrealstock = 5
itemgubun = "10"
stocktype = "real"


'//��ǰ�ڵ� ��ȿ�� �˻�
if itemidArr<>"" then
	dim iA ,arrTemp,arrItemid

    itemidArr = replace(itemidArr,chr(13),"")
	arrTemp = Split(itemidArr,chr(10))

	iA = 0
	do while iA <= ubound(arrTemp)
		if Trim(arrTemp(iA))<>"" and isNumeric(Trim(arrTemp(iA))) then
			arrItemid = arrItemid & Trim(arrTemp(iA)) & ","
		end if
		iA = iA + 1
	loop

	if len(arrItemid)>0 then
		itemidArr = left(arrItemid,len(arrItemid)-1)
	else
		if Not(isNumeric(itemidArr)) then
			itemidArr = ""
		end if
	end if
end if

dim osummarystockbrand
set osummarystockbrand = new CSummaryItemStock
	osummarystockbrand.FPageSize = pagesize
	osummarystockbrand.FCurrPage = page
	osummarystockbrand.FRectItemIdArr = itemidArr
	osummarystockbrand.FRectMakerid = makerid
    osummarystockbrand.FRectMWDiv = mwdiv
    osummarystockbrand.FRectlimitrealstock = limitrealstock

    osummarystockbrand.GetItemSellStateList

%>
<script type="text/javascript" src="/js/barcode.js"></script>
<script type="text/javascript" src="/js/ttpbarcode.js"></script>
<script type="text/javascript" src="/js/DOSHIBAbarcode.js"></script>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script language='javascript'>

function PopItemSellEdit(iitemid){
	var popwin = window.open('/admin/lib/popitemsellinfo.asp?itemid=' + iitemid,'itemselledit','width=500,height=600,scrollbars=yes,resizable=yes')
}

function SubmitSearch() {
	var itemid = document.frm.itemidArr.value;
	 itemid =  itemid.replace(",","\r");    //�޸��� �ٹٲ�ó��
		 for(i=0;i<itemid.length;i++){
			if ( itemid.charCodeAt(i) != "13" && itemid.charCodeAt(i) != "10" && "0123456789".indexOf(itemid.charAt(i)) < 0){
					alert("��ǰ�ڵ�� ���ڸ� �Է°����մϴ�.");
					return;
			}
		}
	frm.action="";
	frm.target="";
    document.frm.submit();
}

</script>

<!-- �˻� ���� -->
<form name="frm" method="get" action="" style="margin:0px;">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="research" value="on">
<input type="hidden" name="page" value="">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="6" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
		<table border="0" cellpadding="5" cellspacing="0" class="a">
			<tr>
				<td>�귣��:	<% drawSelectBoxDesignerwithName "makerid", makerid %></td>
				<td>��ǰ�ڵ�:</td>
				<td ><textarea rows="3" cols="10" name="itemidArr" id="itemidArr"><%=replace(itemidArr,",",chr(10))%></textarea> </td>
			</tr>
		</table>
	</td>
	<td rowspan="6" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="SubmitSearch();">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" height="30">
	<td align="left">
		* ��ǰ����: 10
		&nbsp;&nbsp;
		* �Ǹ� or ��� : ǰ��+�Ͻ�ǰ��
        &nbsp;&nbsp;
        * �ŷ�����:<% drawSelectBoxMWU "mwdiv", mwdiv %>
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" height="30">
	<td align="left">
	    * ��ȿ��� : <input type="text" class="text" size="5" name="limitrealstock" value="<%= limitrealstock %>">
		&nbsp;&nbsp;
		<input type="checkbox" class="checkbox" name="excits" value="Y" checked disabled> 3PL ����
	</td>
</tr>
</table>
</form>
<!-- �˻� �� -->

<p />

<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="37">
		�˻���� : <b><%= osummarystockbrand.FTotalCount %></b>
		&nbsp;
		������ :
		<% if osummarystockbrand.FCurrPage > 1  then %>
			<a href="javascript:GotoPage(<%= page - 1 %>)"><img src="/images/icon_arrow_left.gif" border="0" align="absbottom"></a>
		<% end if %>
		<b><%= page %> / <%= osummarystockbrand.FTotalPage %></b>
		<% if (osummarystockbrand.FTotalpage - osummarystockbrand.FCurrPage)>0  then %>
			<a href="javascript:GotoPage(<%= page + 1 %>)"><img src="/images/icon_arrow_right.gif" border="0" align="absbottom"></a>
		<% end if %>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td><input type="checkbox" name="ckall" onclick="ckAll(this)"></td>
    <td>���ڵ�</td>
    <td>����</td>
	<td>��ǰ�ڵ�</td>
	<td>�ɼ�<br />�ڵ�</td>
	<td>�귣��ID</td>
    <td>��ǰ��<br>[�ɼǸ�]</td>

    <td bgcolor="F4F4F4"><b>�ý���<br>�����</b></td>
	<td bgcolor="F4F4F4"><b>��ȿ<br>���</b></td>

    <td width="30">�Ǹ�<br>����</td>
    <td width="30">���<br>����</td>
    <td width="30">�ɼ�<br />�Ǹ�<br>����</td>
    <td width="30">�ɼ�<br />���<br>����</td>
</tr>
<% if osummarystockbrand.FResultCount>0 then %>
<% for i=0 to osummarystockbrand.FResultCount - 1 %>
<% if osummarystockbrand.FItemList(i).Fisusing="Y" then %>
<tr bgcolor="#FFFFFF" align="center" height="25">
<% else %>
<tr bgcolor="#EEEEEE" align="center" height="25">
<% end if %>
	<form name="frmBuyPrc_<%= i %>" >
	<input type="hidden" id="itembarcode_<%= i %>" name="barcode" value="<%= BF_MakeTenBarcode(osummarystockbrand.FItemList(i).Fitemgubun, osummarystockbrand.FItemList(i).Fitemid, osummarystockbrand.FItemList(i).Fitemoption) %>">
	<input type="hidden" id="publicbarcode_<%= i %>" name="generalbarcode" value="<%= osummarystockbrand.FItemList(i).FpublicBarcode %>">
	<input type="hidden" id="customerprice_<%= i %>" name="orgprice" value="<%= (osummarystockbrand.FItemList(i).Forgprice) %>">
	<input type="hidden" id="itemname_<%= i %>" name="itemname" value="<%= osummarystockbrand.FItemList(i).FItemName %>">
	<input type="hidden" id="itemoptionname_<%= i %>" name="itemoptionname" value="<%= osummarystockbrand.FItemList(i).FItemOptionName %>">
	<input type="hidden" id="sellprice_<%= i %>" name="sellcash" value="<%= osummarystockbrand.FItemList(i).Fsellcash %>">
	<input type="hidden" id="makerid_<%= i %>" name="makerid" value="<%= osummarystockbrand.FItemList(i).FMakerid %>">
	<input type="hidden" id="socname_<%= i %>" name="socname" value="<%= osummarystockbrand.FItemList(i).FMakerid %>">
	<input type="hidden" id="prtidx_<%= i %>" name="prtidx" value="<%= osummarystockbrand.FItemList(i).fprtidx %>">
	<input type="hidden" id="itemrackcode_<%= i %>" name="itemrackcode" value="<%= osummarystockbrand.FItemList(i).fitemrackcode %>">
	<input type="hidden" id="subitemrackcode_<%= i %>" name="subitemrackcode" value="<%= osummarystockbrand.FItemList(i).fsubitemrackcode %>">
	<input type="hidden" name="barcode2" value="<%= BF_MakeTenBarcode(osummarystockbrand.FItemList(i).Fitemgubun, osummarystockbrand.FItemList(i).Fitemid, osummarystockbrand.FItemList(i).Fitemoption) %>_<%= osummarystockbrand.FItemList(i).FpublicBarcode %>">
	<input type="hidden" id="itemgubun_<%= i %>" name="itemgubun" value="<%= osummarystockbrand.FItemList(i).Fitemgubun %>">
	<input type="hidden" id="itemid_<%= i %>" name="itemid" value="<%= osummarystockbrand.FItemList(i).Fitemid %>">
	<input type="hidden" id="itemoption_<%= i %>" name="itemoption" value="<%= osummarystockbrand.FItemList(i).Fitemoption %>">
	<input type="hidden" name="returnitemno" value="<%= osummarystockbrand.FItemList(i).Frealstock*-1 %>">
	<input type="hidden" name="suplycash" value="<%= chkIIF(osummarystockbrand.FItemList(i).IsOffContractExist, osummarystockbrand.FItemList(i).GetOffContractBuycash, osummarystockbrand.FItemList(i).FBuycash) %>">
	<input type="hidden" name="buycash" value="<%= chkIIF(osummarystockbrand.FItemList(i).IsOffContractExist, osummarystockbrand.FItemList(i).GetOffContractBuycash, osummarystockbrand.FItemList(i).FBuycash) %>">
	<input type="hidden" name="mwdiv" value="<%= chkIIF(osummarystockbrand.FItemList(i).IsOffContractExist, osummarystockbrand.FItemList(i).GetOffContractCenterMW, osummarystockbrand.FItemList(i).Fmwdiv) %>">
	<td width=20><input type="checkbox" name="cksel" id="chk_<%= i %>" onClick="AnCheckClick(this);"></td>
    <td><%= osummarystockbrand.FItemList(i).FItemrackcode %></td>
    <td><%= osummarystockbrand.FItemList(i).FItemGubun %></td>
	<td>
	    <% if osummarystockbrand.FItemList(i).FItemGubun="10" then %>
	    <a href="javascript:PopItemSellEdit('<%= osummarystockbrand.FItemList(i).Fitemid %>');"><%= osummarystockbrand.FItemList(i).Fitemid %></a>
	    <% else %>
	    <%= osummarystockbrand.FItemList(i).Fitemid %>
	    <% end if %>
	</td>
    <td><%= osummarystockbrand.FItemList(i).Fitemoption %></td>
	<td><%= osummarystockbrand.FItemList(i).FMakerid %></td>
    <td>
      	<a href="/admin/stock/itemcurrentstock.asp?itemgubun=<%= osummarystockbrand.FItemList(i).FItemGubun %>&itemid=<%= osummarystockbrand.FItemList(i).FItemID %>&itemoption=<%= osummarystockbrand.FItemList(i).FItemOption %>" target=_blank ><%= osummarystockbrand.FItemList(i).Fitemname %></a>
      	<% if osummarystockbrand.FItemList(i).FitemoptionName <>"" then %>
      		<br>
      		<font color="blue">[<%= osummarystockbrand.FItemList(i).FitemoptionName %>]</font>
      	<% end if %>
    </td>

	<td align="right" bgcolor="F4F4F4"><b><%= osummarystockbrand.FItemList(i).Ftotsysstock %></b></td>
	<td align="right" bgcolor="F4F4F4"><b><%= osummarystockbrand.FItemList(i).Frealstock %></td>

	<td><%= fnColor(osummarystockbrand.FItemList(i).Fsellyn,"yn") %></td>
	<td>
		<%= fnColor(osummarystockbrand.FItemList(i).Fisusing,"yn") %>
	</td>
    <td>
		<%= fnColor(osummarystockbrand.FItemList(i).Foptsellyn,"yn") %>
	</td>
    <td>
		<%= fnColor(osummarystockbrand.FItemList(i).Foptisusing,"yn") %>
	</td>
</tr>
</form>
<% next %>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="37" align="center">
		<% if osummarystockbrand.HasPreScroll then %>
		<a href="javascript:NextPage('<%= osummarystockbrand.StartScrollPage-1 %>')">[pre]</a>
		<% else %>
			[pre]
		<% end if %>

		<% for i=0 + osummarystockbrand.StartScrollPage to osummarystockbrand.FScrollCount + osummarystockbrand.StartScrollPage - 1 %>
			<% if i>osummarystockbrand.FTotalpage then Exit for %>
			<% if CStr(page)=CStr(i) then %>
			<font color="red">[<%= i %>]</font>
			<% else %>
			<a href="javascript:NextPage('<%= i %>')">[<%= i %>]</a>
			<% end if %>
		<% next %>

		<% if osummarystockbrand.HasNextScroll then %>
			<a href="javascript:NextPage('<%= i %>')">[next]</a>
		<% else %>
			[next]
		<% end if %>
	</td>
</tr>
<% else %>
    <tr bgcolor="#FFFFFF">
        <td colspan="37" align="center" class="page_link">[�˻������ �����ϴ�.]</td>
    </tr>
<% end if %>
</table>
<%
set osummarystockbrand = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->

<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ����ǰ ���
' Hieditor : 2013.01.15 �̻� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshopitemcls.asp"-->

<%
dim designer,page,usingyn , research, mageview, imageview, itemgubun, itemid, itemname
dim cdl, cdm, cds, i, PriceDiffExists , IsDirectIpchulContractExistsBrand ,publicbarcode
	designer    = RequestCheckVar(request("designer"),32)
	page        = RequestCheckVar(request("page"),9)
	usingyn     = RequestCheckVar(request("usingyn"),1)
	research    = RequestCheckVar(request("research"),9)
	imageview   = RequestCheckVar(request("imageview"),9)
	itemgubun   = RequestCheckVar(request("itemgubun"),16)
	itemid      = RequestCheckVar(request("itemid"),9)
	itemname    = RequestCheckVar(request("itemname"),32)
	publicbarcode    = RequestCheckVar(request("publicbarcode"),20)
	cdl         = RequestCheckVar(request("cdl"),3)
	cdm         = RequestCheckVar(request("cdm"),3)
	cds         = RequestCheckVar(request("cds"),3)
	if page="" then page=1
	if research<>"on" then usingyn="Y"

dim ioffitem
set ioffitem  = new COffShopItem
	ioffitem.FPageSize = 100
	ioffitem.FCurrPage = page
	ioffitem.FRectDesigner = designer
	ioffitem.FRectOnlyUsing = usingyn
	ioffitem.FRectItemgubun = itemgubun
	ioffitem.FRectItemID = itemid
	ioffitem.FRectItemName = html2db(itemname)
	ioffitem.FRectCDL = cdl
	ioffitem.FRectCDM = cdm
	ioffitem.FRectCDS = cds
	ioffitem.FRectpublicbarcode = publicbarcode

	ioffitem.GetOffNOnLineGiftItemList
%>

<script language='javascript'>

//����
function pop_itemedit_gift_edit(ibarcode){
	var pop_itemedit_gift_edit = window.open('/admin/offshop/pop_itemedit_gift_edit.asp?barcode=' + ibarcode,'pop_itemedit_gift_edit','width=1024,height=600,resizable=yes,scrollbars=yes');
	pop_itemedit_gift_edit.focus();
}

//���
function pop_itemedit_gift_new(){
	var pop_itemedit_gift_new;

	pop_itemedit_gift_new = window.open('/admin/offshop/pop_itemedit_gift_edit.asp','pop_itemedit_gift_new','width=1024,height=600,scrollbars=yes,resizable=yes');
	pop_itemedit_gift_new.focus();
}

function ReSearch(page){
	if(frm.itemid.value!=''){
		if (!IsDouble(frm.itemid.value)){
			alert('��ǰ��ȣ�� ���ڸ� �����մϴ�.');
			frm.itemid.focus();
			return;
		}
	}

	frm.page.value = page;
	frm.submit();
}

function GotoPage(page){
    var frm = document.frm;
    frm.page.value = page;
	frm.submit();
}

function jsSelectThisAndCloseWin(itemgubun, itemid, itemoption) {
	opener.ReActWithThis(itemgubun, itemid, itemoption);
	opener.focus();
	window.close();
}

</script>

<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="research" value="on">
<input type="hidden" name="page" value="1">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
		�귣�� :
		<% drawSelectBoxDesignerwithName "designer",designer  %>
		&nbsp;
		����:
		<input type="radio" name="itemgubun" value="" <% if itemgubun = "" then response.write " checked" %>> ��ü
		<input type="radio" name="itemgubun" value="85" <% if itemgubun = "85" then response.write " checked" %>> ON����ǰ
		<input type="radio" name="itemgubun" value="80" <% if itemgubun = "80" then response.write " checked" %>> OFF����ǰ
		&nbsp;
     	���:<% drawSelectBoxUsingYN "usingyn", usingyn %>
	</td>

	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="ReSearch('');">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		��ǰ�ڵ� : <input type="text" class="text" name="itemid" value="<%= itemid %>" size="7" maxlength="9" style="IME-MODE: disabled" />
		&nbsp;
		��ǰ�� : <input type="text" class="text" name="itemname" value="<%= itemname %>" size="24" maxlength="32">
		&nbsp;
		������ڵ� : <input type="text" class="text" name="publicbarcode" value="<%= publicbarcode %>" size="20" maxlength="20">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		<input type="checkbox" name="imageview" value="on" <% if imageview="on" then response.write "checked" %> >�̹�������
	</td>
</tr>
</form>
</table>
<!-- �˻� �� -->

<p>

<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td align="left">
		<input type="button" class="button" value="����ǰ ���" onclick="pop_itemedit_gift_new()">
	</td>
	<td align="right">
	</td>
</tr>
</table>
<!-- �׼� �� -->

<p>

<!-- ����Ʈ ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="25">
		�˻���� : <b><%= ioffitem.FTotalcount %></b>
		<% if ioffitem.FCurrPage > 1  then %>
			<a href="javascript:GotoPage(<%= page - 1 %>)"><img src="/images/icon_arrow_left.gif" border="0" align="absbottom"></a>
		<% end if %>

		<b><%= page %> / <%= ioffitem.FTotalpage %></b>

		<% if (ioffitem.FTotalpage - ioffitem.FCurrPage)>0  then %>
			<a href="javascript:GotoPage(<%= page + 1 %>)"><img src="/images/icon_arrow_right.gif" border="0" align="absbottom"></a>
		<% end if %>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<% if (imageview<>"") then %>
	<td width="50">�̹���</td>
	<% end if %>
	<td>�귣��ID</td>
	<td width="90">��ǰ�ڵ�</td>
	<td>��ǰ��</td>
	<td>�ɼǸ�</td>

	<td width="60">�Һ��ڰ�</td>
	<td width="60">�ǸŰ�</td>

	<td width="60">���԰�</td>
	<td width="60">����<br>���ް�</td>
	<td width="30">����<br>����<br>����</td>

	<td width="30">���<br>����</td>

	<td width="50">���</td>
</tr>
<% for i=0 to ioffitem.FresultCount -1 %>
<% if ioffitem.FItemlist(i).Fisusing="N" then %>
<tr bgcolor="#EEEEEE">
<% else %>
<tr bgcolor="#FFFFFF">
<% end if %>
	<% if (imageview<>"") then %>
		<td width="50" height="55">
			<img src="<%= ioffitem.FItemlist(i).GetImageSmall %>" width=50 height=50 onError="this.src='http://image.10x10.co.kr/images/no_image.gif'" border=0>
		</td>
	<% end if %>
	<td height="30"><%= ioffitem.FItemlist(i).FMakerID %></td>
	<td align="center" >
		<a href="javascript:pop_itemedit_gift_edit('<%= ioffitem.FItemlist(i).GetBarCode %>')" onfocus="this.blur()">
		<%= ioffitem.FItemlist(i).Fitemgubun %>-<%= CHKIIF(ioffitem.FItemlist(i).Fshopitemid>=1000000,Format00(8,ioffitem.FItemlist(i).Fshopitemid),Format00(6,ioffitem.FItemlist(i).Fshopitemid)) %>-<%= ioffitem.FItemlist(i).Fitemoption %>
		</a>
	</td>
	<td>
		<a href="javascript:pop_itemedit_gift_edit('<%= ioffitem.FItemlist(i).GetBarCode %>')" onfocus="this.blur()">
		<%= ioffitem.FItemlist(i).FShopItemName %>
		</a>
	</td>
	<td>
		<%= ioffitem.FItemlist(i).FShopitemOptionname %>
		<% if ioffitem.FItemlist(i).FOnlineOptaddprice<>0 then %>
		    <br>�ɼ��߰��ݾ�: <%= FormatNumber(ioffitem.FItemlist(i).FOnlineOptaddprice,0) %>
		<% end if %>
	</td>
    <td align="right" >
        <%= FormatNumber(ioffitem.FItemlist(i).FShopItemOrgprice, 0) %>
    </td>
	<td align="right" >
		<%= FormatNumber(ioffitem.FItemlist(i).FShopItemprice, 0) %>
	</td>

	<td align="right" >
		<%= FormatNumber(ioffitem.FItemlist(i).Fshopsuplycash, 0) %>
	</td>
	<td align="right" >
		<%= FormatNumber(ioffitem.FItemlist(i).Fshopbuyprice, 0) %>
	</td>
    <td align="center" ><%= ioffitem.FItemlist(i).FCenterMwDiv %></td>
	<td align="center" >
		<%= ioffitem.FItemlist(i).Fisusing %>
	</td>
	<td align="center" >
		<input type="button" class="button" value="����" onclick="jsSelectThisAndCloseWin('<%= ioffitem.FItemlist(i).Fitemgubun %>', '<%= ioffitem.FItemlist(i).Fshopitemid %>', '<%= ioffitem.FItemlist(i).Fitemoption %>')">
	</td>
</tr>
<% next %>
<tr bgcolor="#FFFFFF">
	<td colspan="11" align="center">
	<% if ioffitem.HasPreScroll then %>
		<a href="javascript:ReSearch('<%= ioffitem.StartScrollPage-1 %>')">[pre]</a>
	<% else %>
		[pre]
	<% end if %>

	<% for i=0 + ioffitem.StartScrollPage to ioffitem.FScrollCount + ioffitem.StartScrollPage - 1 %>
		<% if i>ioffitem.FTotalpage then Exit for %>
		<% if CStr(page)=CStr(i) then %>
		<font color="red">[<%= i %>]</font>
		<% else %>
		<a href="javascript:ReSearch('<%= i %>')">[<%= i %>]</a>
		<% end if %>
	<% next %>

	<% if ioffitem.HasNextScroll then %>
		<a href="javascript:ReSearch('<%= i %>')">[next]</a>
	<% else %>
		[next]
	<% end if %>
	</td>
</tr>
</table>

<%
set ioffitem  = Nothing
%>
<!-- #include virtual="/common/lib/commonbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
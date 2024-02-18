<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/PlusSaleCls.asp"-->
<%

dim benefit_group_no
dim i, j, k
dim page :page = 1

benefit_group_no = request("benefit_group_no")
if (benefit_group_no <> "") then
    if Not IsNumeric(benefit_group_no) then
        benefit_group_no = ""
    end if
end if

if (benefit_group_no = "") then
    response.write "�߸��� �����Դϴ�."
    dbget.close : response.end
end if


'// ===============================================
'// ���ñ׷�
'// ===============================================
dim oCBuyBenefitGroupItem
set oCBuyBenefitGroupItem = new CBuyBenefit
oCBuyBenefitGroupItem.FRectBenefitGroupNo = CHKIIF(benefit_group_no="", "-1", benefit_group_no)
''oCBuyBenefitGroupItem.FRectUseYN = "Y"
oCBuyBenefitGroupItem.FPageSize = 100
oCBuyBenefitGroupItem.FCurrPage = page

oCBuyBenefitGroupItem.GetBuyBenefitGroupItemList

%>
<script language='javascript'>

function jsItemModify(benefit_group_no, idx) {
    var popwin = window.open("PlusSaleItemModify.asp?benefit_group_no=<%= benefit_group_no %>&plus_sale_item_idx=" + idx,"jsItemModify" + idx,"width=1200 height=600 scrollbars=yes resizable=yes");
    popwin.focus();
}

</script>
<!-- ����Ʈ ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="17">
        <img src="/images/icon_arrow_down.gif" align="absbottom">
        <font color="red"><strong>���ñ׷��ǰ ���</strong></font>
		�˻���� : <b><%= oCBuyBenefitGroupItem.FTotalCount %></b>
		&nbsp;
		������ : <b><%= page %> / <%= oCBuyBenefitGroupItem.FTotalPage %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="25">
	<td width=40 rowspan="2">IDX</td>
	<td width=80 rowspan="2">��ǰ�ڵ�</td>
    <td rowspan="2">��ǰ��</td>
    <td colspan="4">�÷���</td>
    <td width=50 rowspan="2">����<br />����</td>
    <td width=50 rowspan="2">����<br />����</td>
    <td width=50 rowspan="2">�ִ�<br />���ż�</td>
    <td rowspan="2">��������</td>
    <td rowspan="2">���ǻ���</td>
    <td width=40 rowspan="2">����<br />��ȣ</td>
    <td width=40 rowspan="2">�Ǹ�<br />����</td>
    <td width=40 rowspan="2">�ɼ�<br />����</td>
    <td width=40 rowspan="2">���<br />����</td>
    <td rowspan="2">���</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="25">
    <td width=60>�Һ��ڰ�</td>
    <td width=60>���ΰ���</td>
    <td width=60>������</td>
    <td width=60>���θ��԰�</td>
</tr>
<% for i=0 to oCBuyBenefitGroupItem.FResultcount-1 %>
<tr bgcolor="<%= CHKIIF(oCBuyBenefitGroupItem.FItemList(i).Fuse_yn="Y", "#FFFFFF", "#EEEEEE") %>" align="center" height="25">
    <td><%= oCBuyBenefitGroupItem.FItemList(i).Fplus_sale_item_idx %></td>
    <td><%= oCBuyBenefitGroupItem.FItemList(i).Fitemid %></td>
    <td><%= oCBuyBenefitGroupItem.FItemList(i).Fitemname %></td>
    <td><%= FormatNumber(oCBuyBenefitGroupItem.FItemList(i).Fsellcash, 0) %></td>
    <td><%= FormatNumber(oCBuyBenefitGroupItem.FItemList(i).Fplus_sale_price, 0) %></td>
    <td><%= oCBuyBenefitGroupItem.FItemList(i).Fplus_sale_pct %> %</td>
    <td><%= FormatNumber(oCBuyBenefitGroupItem.FItemList(i).Fplus_sale_buyprice, 0) %></td>
    <td><%= oCBuyBenefitGroupItem.FItemList(i).Flimit_yn %></td>
    <td><%= FormatNumber(oCBuyBenefitGroupItem.FItemList(i).Flimit_cnt, 0) %></td>
    <td><%= oCBuyBenefitGroupItem.FItemList(i).Fmax_buy_cnt %></td>
    <td><%= oCBuyBenefitGroupItem.FItemList(i).Fbadge_contents %></td>
    <td><%= oCBuyBenefitGroupItem.FItemList(i).Fnotice %></td>
    <td><%= oCBuyBenefitGroupItem.FItemList(i).Fsort_no %></td>
    <td><%= oCBuyBenefitGroupItem.FItemList(i).Fsell_cnt %></td>
    <td><%= oCBuyBenefitGroupItem.FItemList(i).Fopt_cnt %></td>
    <td><%= oCBuyBenefitGroupItem.FItemList(i).Fuse_yn %></td>
    <td>
        <input type="button" class="button" value="��ǰ����" onclick="jsItemModify(<%= oCBuyBenefitGroupItem.FItemList(i).Fbenefit_group_no %>, <%= oCBuyBenefitGroupItem.FItemList(i).Fplus_sale_item_idx %>)">
    </td>
</tr>
<% next %>
</table>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
    <tr valign="bottom" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="center">
            &nbsp;
        	<input type="button" class="button" value=" ��ǰ�߰� " onclick="jsItemModify(<%= benefit_group_no %>, '')">
	    </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    <tr valign="top" height="10">
        <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_08.gif"></td>
        <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
    </tr>
</table>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->

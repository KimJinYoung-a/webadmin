<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stock/summary_itemstockcls.asp"-->
<%
dim makerid, mwdiv, isusing, itemid, cate_large
dim page, research, finType

makerid = RequestCheckVar(request("makerid"),32)
itemid  = RequestCheckVar(request("itemid"),9)
isusing = RequestCheckVar(request("isusing"),9)
mwdiv   = RequestCheckVar(request("mwdiv"),9)
page    = RequestCheckVar(request("page"),9)
research= RequestCheckVar(request("research"),9)
cate_large = RequestCheckVar(request("cate_large"),3)
finType = RequestCheckVar(request("finType"),9)

if (page="") then page=1
if (research="") then
    if (isusing="") then isusing="Y"
    if (mwdiv="") then mwdiv="MW"

    if (finType="") then finType="on"
end if

dim oStatSList
set oStatSList = new CSummaryItemStock
oStatSList.FCurrPage    = page
oStatSList.FPageSize        = 30
oStatSList.FRectCd1         = cate_large
oStatSList.FRectMakerid     = makerid
oStatSList.FRectItemID      = itemid
oStatSList.FRectOnlyIsUsing = isusing
oStatSList.FRectMWDiv       = mwdiv
oStatSList.FRectState       = finType

oStatSList.GetImsiSoldOutList

dim i
%>
<script language='javascript'>
function popDanjongSet(iitemid, itemoption, actType){
    var popwin = window.open('/common/popitemdanjongSet.asp?itemid=' + iitemid + '&itemoption=' + itemoption + '&actType=' + actType,'popitemdanjongSet','width=900, height=400, scrollbars=yes, resizable=yes');
	popwin.focus();
}

function PopItemSellEdit(iitemid){
	var popwin = window.open('/common/pop_simpleitemedit.asp?itemid=' + iitemid,'itemselledit','width=500,height=600,scrollbars=yes,resizable=yes')
	popwin.focus();
}

function PopItemDetail(itemid, itemoption){
	var popwin = window.open('/admin/stock/itemcurrentstock.asp?itemid=' + itemid + '&itemoption=' + itemoption,'popitemdetail','width=1000, height=600, scrollbars=yes');
	popwin.focus();
}

function NextPage(page){
    document.frm.page.value=page;
    document.frm.submit();

}

</script>

<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="#999999">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="research" value="on">
	<input type="hidden" name="page" value="1">
	<tr align="center" bgcolor="#FFFFFF" >
	    <td rowspan="2" width="50" bgcolor="#EEEEEE">�˻�<br>����</td>
        <td align="left">
        	�귣��: <% drawSelectBoxDesignerwithName "makerid",makerid %>&nbsp;
        	&nbsp;
		    ī�װ� : <% SelectBoxBrandCategory "cate_large", cate_large %>
        	�ŷ�����: <% drawSelectBoxMWU "mwdiv",mwdiv %>&nbsp;
        	<br>
        	��ǰ�ڵ�: <input type="text" name="itemid" value="<%= itemid %>" size="9" maxlength="9">
        	&nbsp;
        	<input type="checkbox" name="finType" <%= ChkIIF(finType="on","checked","") %> >(����/���԰� ����)��ó�� ������

        	<input type=checkbox name="isusing" value="on" <% if isusing="on" then response.write "checked" %> >����ǰ��
        	<br>
        </td>
        <td rowspan="2" width="50" bgcolor="#EEEEEE">
        	<a href="javascript:document.frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a><br>
        </td>
	</tr>
	</form>
</table>
<!-- ǥ ��ܹ� ��-->

<!-- �׼� ���� -->
<!--
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<tr>
		<td align="right">
			<input type="button" class="button" value="��ü����" onClick="">
			&nbsp;
			������ : <input type="text" class="text" name="" size="3" maxlength="5">
			<input type="button" class="button" value="���û�ǰ����" onClick="">
			&nbsp;
			<input type="button" class="button" value="����" onClick="">
		</td>

	</tr>
</table>
-->
<!-- �׼� �� -->
<p>
<!-- ����Ʈ ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="#999999">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="25">
			�˻���� : <b><%= oStatSList.FTotalCount %></b>
			&nbsp;
			������ : <b><%= page %> / <%= oStatSList.FTotalPage %></b>
		</td>
	</tr>
	<tr align="center" bgcolor="#E6E6E6">
    	<!-- <td>����</td> -->
		<td width="50">�̹���</td>
		<td width="70">�귣��</td>
		<td width="40">��ǰ<br>�ڵ�</td>
		<td>��ǰ��<br>(�ɼǸ�)</td>
		<td width="35">���<br>����</td>
        <td width="25">��ü<br>�԰�<br>��ǰ</td>
        <td width="25">��ü<br>�Ǹ�<br>��ǰ</td>
        <td width="25">��ü<br>���<br>��ǰ</td>
        <td width="25">��Ÿ<br>���<br>��ǰ</td>
        <td width="25">CS<br>���<br>��ǰ</td>

		<td width="50">��<br>�ǻ�<br>����</td>
		<td width="50">�ǻ�<br>���</td>
		<td width="50">��<br>�ҷ�</td>
		<td width="50">��ȿ<br>���</td>

		<!--
		<td width="25">��<br>�ҷ�</td>
        <td width="25">��<br>�ǻ�<br>����</td>
        <td width="25">�ǻ�<br>���</td>
		-->

        <td width="25">��<br>��ǰ<br>�غ�</td>
        <td width="25">���<br>�ľ�<br>���</td>
        <td width="25">ON<br>����<br>�Ϸ�</td>
        <td width="25">ON<br>�ֹ�<br>����</td>
        <td width="25">����<br>��<br>���</td>
		<td width="40">�Ǹ�<br>����</td>
		<td width="50">����<br>����</td>
        <td width="35">����<br>����</td>
        <td width="60">���԰�<br>������<br>(���)</td>
        <td width="35">����<br>ó��</td>
    </tr>
<% for i=0 to oStatSList.FresultCount-1 %>
    <% if oStatSList.FItemList(i).Fisusing="Y" then %>
    <tr bgcolor="#FFFFFF" align="center">
    <% else %>
    <tr bgcolor="#EEEEEE" align="center">
    <% end if %>
    	<!-- <td><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"></td> -->
		<td><img src="<%= oStatSList.FItemList(i).Fimgsmall %>" width="50" height="50"></td>
		<td align="left">
          <%= oStatSList.FItemList(i).FMakerID %>
        </td>
		<td>
          <a href="javascript:PopItemSellEdit('<%= oStatSList.FItemList(i).FItemID %>');"><%= oStatSList.FItemList(i).FItemID %></a>
        </td>
		<td align="left">
          <a href="javascript:PopItemDetail('<%= oStatSList.FItemList(i).FItemID %>','<%= oStatSList.FItemList(i).FItemOption %>')"><%= oStatSList.FItemList(i).FItemName %></a>
        <% if (oStatSList.FItemList(i).FItemOptionName <> "") then %>
          <br>(<font color="#3333CC"><%= oStatSList.FItemList(i).FItemOptionName %></font>)
        <% end if %>
        </td>
        <td><font color="<%= mwdivColor(oStatSList.FItemList(i).Fmwdiv) %>"><%= mwdivName(oStatSList.FItemList(i).Fmwdiv) %></font></td>
		<td><%= oStatSList.FItemList(i).Ftotipgono %></td>
		<td><%= -1*oStatSList.FItemList(i).Ftotsellno %></td>
		<td><%= oStatSList.FItemList(i).Foffchulgono + oStatSList.FItemList(i).Foffrechulgono %></td>
        <td><%= oStatSList.FItemList(i).Fetcchulgono + oStatSList.FItemList(i).Fetcrechulgono %></td>
        <td><%= oStatSList.FItemList(i).Ferrcsno %></td>

		<td><%= oStatSList.FItemList(i).Ferrrealcheckno %></td>
		<td><%= oStatSList.FItemList(i).getErrAssignStock %></td>
		<td><%= oStatSList.FItemList(i).Ferrbaditemno %></td>
		<td><%= oStatSList.FItemList(i).Frealstock %></td>

		<!--
        <td><%= oStatSList.FItemList(i).Ferrbaditemno %></td>
        <td><%= oStatSList.FItemList(i).Ferrrealcheckno %></td>
        <td><b><%= oStatSList.FItemList(i).Frealstock %></b></td>
		-->

		<td><%= oStatSList.FItemList(i).Fipkumdiv5 + oStatSList.FItemList(i).Foffconfirmno %></td>
        <td><b><%= oStatSList.FItemList(i).GetCheckStockNo %></b></td>
        <td><%= oStatSList.FItemList(i).Fipkumdiv4 %></td>
        <td><%= oStatSList.FItemList(i).Fipkumdiv2 %></td>
        <td><b><%= oStatSList.FItemList(i).GetLimitStockNo %></b></td>
        <td>
        	<%= oStatSList.FItemList(i).Fsellyn %>
        </td>

        <td>
        <% if (oStatSList.FItemList(i).Flimityn = "Y") then %>
          	����(<%= oStatSList.FItemList(i).GetLimitStr %>)
            <% if (oStatSList.FItemList(i).Foptlimityn = "Y") then %>
            <br>(<%= oStatSList.FItemList(i).Foptlimitno %>/<%= oStatSList.FItemList(i).Foptlimitsold %>)
            <% else %>
            <br>(<%= oStatSList.FItemList(i).FLimitNo %>/<%= oStatSList.FItemList(i).FLimitSold %>)
          	<% end if %>
        <% end if %>
        </td>
        <td><%= oStatSList.FItemList(i).getDanjongNameHTML %></td>
        <td>
            <% if (Not IsNull(oStatSList.FItemList(i).Fstockreipgodate)) then %>
            <a href="javascript:popDanjongSet('<%= oStatSList.FItemList(i).FItemID %>','<%= oStatSList.FItemList(i).FItemOption %>','R');"><%= oStatSList.FItemList(i).Fstockreipgodate %></a>
            <% else %>
            <a href="javascript:popDanjongSet('<%= oStatSList.FItemList(i).FItemID %>','<%= oStatSList.FItemList(i).FItemOption %>','R');"><img src="/images/icon_arrow_link.gif" width="14" border="0"></a>
            <% end if %>
        </td>
        <td>
            <% if (oStatSList.FItemList(i).FDanjongyn<>"M") and (oStatSList.FItemList(i).FDanjongyn<>"Y") then %>
            <a href="javascript:popDanjongSet('<%= oStatSList.FItemList(i).FItemID %>','<%= oStatSList.FItemList(i).FItemOption %>','D');"><img src="/images/icon_arrow_link.gif" width="14" border="0"></a>
            <% end if %>
        </td>

	</tr>
<% next %>
    <tr height="25" bgcolor="FFFFFF">
		<td colspan="25" align="center">
		<% if oStatSList.HasPreScroll then %>
    		<a href="javascript:NextPage('<%= oStatSList.StartScrollPage-1 %>');">[pre]</a>
    	<% else %>
    		[pre]
    	<% end if %>

    	<% for i=0 + oStatSList.StartScrollPage to oStatSList.FScrollCount + oStatSList.StartScrollPage - 1 %>
    		<% if i>oStatSList.FTotalpage then Exit for %>
    		<% if CStr(page)=CStr(i) then %>
    		<font color="red">[<%= i %>]</font>
    		<% else %>
    		<a href="javascript:NextPage('<%= i %>');">[<%= i %>]</a>
    		<% end if %>
    	<% next %>

    	<% if oStatSList.HasNextScroll then %>
    		<a href="javascript:NextPage('<%= i %>');">[next]</a>
    	<% else %>
    		[next]
    	<% end if %>
	    </td>
	</tr>
</table>

<%
set oStatSList = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->

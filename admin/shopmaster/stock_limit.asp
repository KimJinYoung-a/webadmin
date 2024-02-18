<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : [����]��������
' History : 		   �̻� ����
'			2016.03.29 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stock/summary_itemstockcls.asp"-->
<%
dim shopid, diff, isusing, diffdiv, mwdiv, orderby , currPage, OnlySellyn, BasicMonth, i, itemid
dim searchtype, rackcode2, fromrackcode2, torackcode2
dim excits
	shopid  		= request("shopid")
	diff    		= trim(request("diff"))
	isusing 		= requestcheckvar(request("isusing"),1)
	diffdiv 		= trim(request("diffdiv"))
	OnlySellyn 		= request("OnlySellyn")
	mwdiv   		= request("mwdiv")
	orderby 		= request("orderby")
	currPage 		= getNumeric(requestcheckvar(request("cp"),10))
	itemid 			= getNumeric(requestcheckvar(request("itemid"),10))
	searchtype  	= requestCheckvar(request("searchtype"),1)
	rackcode2   	= requestCheckvar(request("rackcode2"),2)
	fromrackcode2  	= requestCheckvar(request("fromrackcode2"),2)
	torackcode2  	= requestCheckvar(request("torackcode2"),2)
	excits  		= requestCheckvar(request("excits"),2)

IF currPage="" Then currPage = 1
if (diffdiv = "") then
	diffdiv = "percent"
end if

if (diff = "") then
	diff = "30"
end if

if (diff < 0) then
	diff = -1 * diff
end if

if searchtype="" then searchtype = "F"
if ((request("research") = "") and (isusing = "")) then
	isusing = "Y"
end if

if ((request("research") = "") and (OnlySellyn = "")) then
	OnlySellyn = "YS"
end if

if ((request("research") = "") and (mwdiv = "")) then
	mwdiv = "MW"
end if

if (request("research") = "") then
	excits = "Y"
end if


BasicMonth = Left(CStr(DateSerial(Year(now()),Month(now())-1,1)),7)

dim osummarystock
set osummarystock = new CSummaryItemStock
	osummarystock.FCurrPage = currPage
	osummarystock.FPageSize    = 100
	osummarystock.FRectMakerid = shopid
	osummarystock.FRectParameter = diff
	osummarystock.FRectDiffDiv = diffdiv
	osummarystock.FRectOnlyIsUsing = isusing
	osummarystock.FRectOnlySellyn = OnlySellyn
	osummarystock.FRectMwDiv      = mwdiv
	osummarystock.FRectOrderBy      = orderby
	osummarystock.FRectitemid = itemid
	osummarystock.FRectSearchType = searchtype
	osummarystock.FRectRackCode = rackcode2
	osummarystock.FRectFromRackcode2 = fromrackcode2
	osummarystock.FRectToRackcode2 = torackcode2
	osummarystock.FRectExcIts = excits

	osummarystock.GetCurrentStockByOnlineBrandLimit

%>

<script language='javascript'>

function PopItemSellEdit(iitemid){
	var popwin = window.open('/admin/lib/popitemsellinfo.asp?itemid=' + iitemid,'itemselledit','width=500 height=600')
}

function PopItemDetail(itemid, itemoption){
	var popwin = window.open('/admin/stock/itemcurrentstock.asp?itemid=' + itemid + '&itemoption=' + itemoption,'popitemdetail','width=1000, height=600, scrollbars=yes');
	popwin.focus();
}

function NextPage(v){
	if(frm.itemid.value!=''){
		if (!IsDouble(frm.itemid.value)){
			alert('��ǰ�ڵ�� ���ڸ� �����մϴ�.');
			frm.itemid.focus();
			return;
		}
	}

	document.frm.cp.value=v;
	document.frm.submit();
}

//�����Է�
function popRealErrInput(itemgubun,itemid,itemoption){
	var popwin = window.open('/common/poprealerrinput.asp?itemgubun=' + itemgubun + '&itemid=' + itemid + '&itemoption=' + itemoption + '&BasicMonth=<%= BasicMonth %>','poprealerrinput','width=1024,height=768,scrollbar=yes,resizable=yes')
	popwin.focus();
}

</script>

<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="research" value="on">
<input type="hidden" name="cp" value="<%= currPage %>">
<tr align="center" bgcolor="#FFFFFF" height="30">
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
		�귣�� : <% drawSelectBoxDesignerwithName "shopid",shopid %>
		&nbsp;&nbsp;
		<input type="radio" name="diffdiv" value="over" <% if (diffdiv = "over") then %>checked<% end if %>> �����ʰ�&nbsp;&nbsp;
		<input type="radio" name="diffdiv" value="number" <% if (diffdiv = "number") then %>checked<% end if %>> ����&nbsp;&nbsp;
      	<input type="radio" name="diffdiv" value="percent" <% if (diffdiv = "percent") then %>checked<% end if %>> �ۼ�Ʈ&nbsp;&nbsp;
    	���� : <input type="text" class="text" name="diff" value="<%= diff %>" size="6">
    	&nbsp;&nbsp;
		<% if (diffdiv = "over") then %>
			* ������ ���� �������������� ���� ��� �Դϴ�.
		<% elseif (diffdiv = "number") then %>
			* ���������� ���������������� ���̰� <strong><%= diff %>��</strong> �ʰ��Ǵ� ��ǰ�Դϴ�.
		<% else %>
			* ���������� ���������������� ���̰� <strong><%= diff %>%</strong> �ʰ��Ǵ� ��ǰ�Դϴ�.
		<% end if %>
	</td>
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="NextPage('1');">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" height="30">
	<td align="left">
		�Ǹ� : <% drawSelectBoxSellYN "OnlySellyn", OnlySellyn %>
     	&nbsp;&nbsp;
     	��� : <% drawSelectBoxUsingYN "isusing", isusing %>
     	&nbsp;&nbsp;
     	�ŷ����� : <% drawSelectBoxMWU "mwdiv", mwdiv %>
     	&nbsp;&nbsp;
		���� :
		<select class="select" name="orderby">
			<option  value="">�ǻ����</option> <!-- �ʱⰪ -->
			<option  value="makerid" <%= ChkIIF(orderby="makerid","selected","") %> >�귣��ID</option> <!-- ���ĺ����� -->
			<option  value="itemrackcode" <%= ChkIIF(orderby="itemrackcode","selected","") %> >��ǰ���ڵ�</option> <!-- �������ں��� -->
			<option  value="itemid" <%= ChkIIF(orderby="itemid","selected","") %> >��ǰ�ڵ�</option>
		</select>
		&nbsp;&nbsp;
		��ǰ�ڵ� : <input type="text" class="text" name="itemid" value="<%= itemid %>" size="10">
		&nbsp;&nbsp;
		���ڵ� :
		<input type="radio" name="searchtype" value="F" <% if (searchtype = "F") then %>checked<% end if %> >
		<input type="text" name=rackcode2 value="<%= rackcode2 %>" maxlength="2" size="2" class="text"> (�� 2�ڸ�)
		&nbsp;
		<input type="radio" name="searchtype" value="R" <% if (searchtype = "R") then %>checked<% end if %> >
    	<input type="text" name=fromrackcode2 value="<%= fromrackcode2 %>" maxlength="2" size="2" class="text">
		~
		<input type="text" name=torackcode2 value="<%= torackcode2 %>" maxlength="2" size="2" class="text"> (�� 2�ڸ�)
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" height="30">
	<td align="left">
		<input type="checkbox" class="checkbox" name="excits" value="Y" <%= CHKIIF(excits="Y", "checked", "") %> > ���̶�� ����
	</td>
</tr>
</form>
</table>

<br>
<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:1;">
<tr>
	<td align="left"></td>
	<td align="right"></td>
</tr>
</tr>
</table>
<!-- �׼� �� -->

<!-- ����Ʈ ���� -->
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="25">
		�˻���� : <b><%= osummarystock.FresultCount %></b>
		&nbsp;
		������ : <b><%= currPage %>/ <%= osummarystock.FTotalPage %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="80">�귣��</td>
	<td width="50">�̹���</td>
    <td width="40">��ǰ<br>���ڵ�</td>
	<td width="40">��ǰ<br>�ڵ�</td>
	<td width="40">�ɼ�<br>�ڵ�</td>
	<td>��ǰ��<br>(�ɼǸ�)</td>
	<td width="35">���<br>����</td>
    <td width="35">��ü<br>�԰�<br>��ǰ</td>
    <td width="35">��ü<br>�Ǹ�<br>��ǰ</td>
    <td width="35">��ü<br>���<br>��ǰ</td>
    <td width="35">��Ÿ<br>���<br>��ǰ</td>
	<td width="50">��<br>�ǻ�<br>����</td>
	<td width="50">�ǻ�<br>���</td>
	<td width="50">��<br>�ҷ�</td>
	<td width="50">��ȿ<br>���</td>
	<!--<td width="30">��<br>�ҷ�</td>
    <td width="30">��<br>�ǻ�<br>����</td>
    <td width="35">�ǻ�<br>���</td>-->
	<td width="30">��<br>��ǰ<br>�غ�</td>
    <td width="35">���<br>�ľ�<br>���</td>
    <td width="30">ON<br>����<br>�Ϸ�</td>
    <td width="30">ON<br>�ֹ�<br>����</td>
    <td width="35">����<br>��<br>���</td>
    <td width="35">����</td>
	<td width="35">����<br>����</td>
	<td width="35">�Ǹ�<br>����</td>
	<td width="50">����<br>�Է�</td>
</tr>
<% if osummarystock.FresultCount > 0 then %>
	<% for i=0 to osummarystock.FresultCount - 1 %>
	<% if osummarystock.FItemList(i).Fisusing="Y" and (osummarystock.FItemList(i).GetLimitStr <> "") then %>
		<tr bgcolor="#FFFFFF" align="center">
	<% elseif IsNull(osummarystock.FItemList(i).GetLimitStr) then  %>
		<tr bgcolor="#FF0000" align="center">
	<% else %>
		<tr bgcolor="#EEEEEE" align="center">
	<% end if %>

		<td><%= osummarystock.FItemList(i).FMakerID %></td>
		<td><img src="<%= osummarystock.FItemList(i).Fimgsmall %>" width="50" height="50"></td>
        <td><%= osummarystock.FItemList(i).Fitemrackcode %></td>
		<td><a href="javascript:PopItemSellEdit('<%= osummarystock.FItemList(i).FItemID %>');"><%= osummarystock.FItemList(i).FItemID %></a></td>
		<td <%=CHKIIF(osummarystock.FItemList(i).Foptioncnt>0 and osummarystock.FItemList(i).FItemOption="0000"," bgcolor='#FF3333'","")%>><%= osummarystock.FItemList(i).FItemOption %></td>
		<td align="left">
	      	<a href="javascript:PopItemDetail('<%= osummarystock.FItemList(i).FItemID %>','<%= osummarystock.FItemList(i).FItemOption %>')"><%= osummarystock.FItemList(i).FItemName %></a>
	    	<% if (osummarystock.FItemList(i).FItemOptionName <> "") then %>
	      	<br><font color="blue">[<%= osummarystock.FItemList(i).FItemOptionName %>]</font>
	    	<% end if %>
	    </td>
	    <td><%= fnColor(osummarystock.FItemList(i).Fmwdiv,"mw") %></td>
		<td><%= osummarystock.FItemList(i).Ftotipgono %></td>
		<td><%= -1*osummarystock.FItemList(i).Ftotsellno %></td>
		<td><%= osummarystock.FItemList(i).Foffchulgono + osummarystock.FItemList(i).Foffrechulgono %></td>
	    <td><%= osummarystock.FItemList(i).Fetcchulgono + osummarystock.FItemList(i).Fetcrechulgono %></td>
		<td align="right"><b><%= FormatNumber(osummarystock.FItemList(i).Ferrrealcheckno, 0) %></b>&nbsp;</td>
	    <td align="right"><%= FormatNumber(osummarystock.FItemList(i).getErrAssignStock, 0) %>&nbsp;</td>
		<td align="right">
			<%= FormatNumber(osummarystock.FItemList(i).Ferrbaditemno, 0) %>&nbsp;
		</td>
	    <td align="right"><%= FormatNumber(osummarystock.FItemList(i).Frealstock, 0) %>&nbsp;</td>
		<!--<td><%= osummarystock.FItemList(i).Ferrbaditemno %></td>
	    <td><%= osummarystock.FItemList(i).Ferrrealcheckno %></td>
	    <td><b><%= osummarystock.FItemList(i).Frealstock %></b></td>-->
		<td><%= osummarystock.FItemList(i).Fipkumdiv5 + osummarystock.FItemList(i).Foffconfirmno %></td>
	    <td><b><%= osummarystock.FItemList(i).GetCheckStockNo %></b></td>
	    <td><%= osummarystock.FItemList(i).Fipkumdiv4 %></td>
	    <td><%= osummarystock.FItemList(i).Fipkumdiv2 %></td>
	    <td><b><%= osummarystock.FItemList(i).GetLimitStockNo %></b></td>
	   		<% if (diffdiv = "number") or (diffdiv = "over") then %>
	    <td><%= osummarystock.FItemList(i).GetLimitStockNo - osummarystock.FItemList(i).GetLimitStr %></td>
		<% elseif (diffdiv = "percent") then %>
	    <td><%= round((100 - (osummarystock.FItemList(i).GetLimitStr * 100 / osummarystock.FItemList(i).GetLimitStockNo)),1) %>%</td>
	    	<% end if %>
	    <td>
	      	����<br>
	      	(<%= osummarystock.FItemList(i).GetLimitStr %>)

			<% 'if (osummarystock.FItemList(i).Foptlimityn = "Y") then %>
				<!--<br>(<% '= osummarystock.FItemList(i).Foptlimitno %>/<% '= osummarystock.FItemList(i).Foptlimitsold %>)-->
			<% 'else %>
				<!--<br>(<% '= osummarystock.FItemList(i).FLimitNo %>/<% '= osummarystock.FItemList(i).FLimitSold %>)-->
			<% 'end if %>
	    </td>
	    <td><%= fnColor(osummarystock.FItemList(i).Fsellyn,"yn") %></td>
		<td>
			<input type="button" class="button" value="����" onclick="popRealErrInput('<%= osummarystock.FItemList(i).Fitemgubun %>','<%= osummarystock.FItemList(i).Fitemid %>','<%= osummarystock.FItemList(i).Fitemoption %>');">
		</td>
	</tr>
	<% next %>

	<tr height="25" bgcolor="FFFFFF">
		<td colspan="25" align="center">
			<% if osummarystock.HasPreScroll then %>
	    		<a href="javascript:NextPage('<%= osummarystock.StartScrollPage-1 %>')">[pre]</a>
	    	<% else %>
	    		[pre]
	    	<% end if %>

	    	<% for i=0 + osummarystock.StartScrollPage to osummarystock.FScrollCount + osummarystock.StartScrollPage - 1 %>
	    		<% if i>osummarystock.FTotalpage then Exit for %>
	    		<% if CStr(currPage)=CStr(i) then %>
	    		<font color="red">[<%= i %>]</font>
	    		<% else %>
	    		<a href="javascript:NextPage('<%= i %>')">[<%= i %>]</a>
	    		<% end if %>
	    	<% next %>

	    	<% if osummarystock.HasNextScroll then %>
	    		<a href="javascript:NextPage('<%= i %>')">[next]</a>
	    	<% else %>
	    		[next]
	    	<% end if %>
		</td>
	</tr>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="25" align="center" class="page_link">[�˻������ �����ϴ�.]</td>
	</tr>
<% end if %>
</table>

<form name="frmArrupdate" method="post" action="dolimitsoldset.asp">
<input type="hidden" name="mode" value="arr">
<input type="hidden" name="itemid" value="">
<input type="hidden" name="dispyn" value="">
<input type="hidden" name="sellyn" value="">
</form>

<%
set osummarystock = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->

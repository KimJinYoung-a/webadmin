<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/designer/lib/designerbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/items/itemcls_v2.asp"-->
<%

dim itemid, makerid, itemname
dim sellyn, isusing, danjongyn, limityn, mwdiv
dim page
dim searchtype
dim showimage

itemid  = RequestCheckVar(request("itemid"),10)
makerid = RequestCheckVar(request("makerid"),32)
itemname = RequestCheckVar(request("itemname"),32)

sellyn  = RequestCheckVar(request("sellyn"),10)
isusing = RequestCheckVar(request("isusing"),10)
danjongyn = RequestCheckVar(request("danjongyn"),10)
limityn = RequestCheckVar(request("limityn"),10)
mwdiv = RequestCheckVar(request("mwdiv"),10)

page = RequestCheckVar(request("page"),10)

searchtype = RequestCheckVar(request("searchtype"),32)

showimage = RequestCheckVar(request("showimage"),32)



if (sellyn="") then sellyn="A"

if (page="") then page=1

''if (isusing="") then isusing="Y"
''����ϴ� ��ǰ�� ǥ�÷� ����
isusing="Y"

'��ǰ�ڵ� ��ȿ�� �˻�(2008.08.01;������)
if itemid<>"" then
	if Not(isNumeric(itemid)) then
		Response.Write "<script language=javascript>alert('[" & itemid & "]��(��) ��ȿ�� ��ǰ�ڵ尡 �ƴմϴ�.');history.back();</script>"
		dbget.close()	:	response.End
	end if
end if

'==============================================================================
dim oitem

set oitem = new CItem

oitem.FRectMakerId = session("ssBctID")
oitem.FRectItemid = itemid
oitem.FRectItemName = itemname

oitem.FRectSearchType = searchtype

if (showimage = "Y") then
	oitem.GetJupsuProductList_CS
else
	oitem.GetJupsuProductListQuick_CS
end if

dim i

dim jupsuChulgoSUM, confirmChulgoSUM, jupsuReturnSUM

%>
<script language='javascript'>
function NextPage(ipage){
	document.frm.page.value= ipage;
	SubmitSearch();
}

function SubmitSearch(){
	if ((document.frm.itemid.value != "") && ((document.frm.itemid.value*0) != 0)) {
	    alert("��ǰ�ڵ忡�� ���ڸ� �Է��� �����մϴ�.");
	    document.frm.itemid.focus();
	    return;
    }
	document.frm.submit();
}

function xlDown(){
	xlfrm.target="iiframeXL";
	xlfrm.action="upchecs_jupsu_item_list_XL.asp";
	xlfrm.submit();
}
</script>


<!-- ǥ ��ܹ� ����-->

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method=get>
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="page" >
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
		<td align="left">
			��ǰ�ڵ� :
			<input type="text" class="text" name="itemid" value="<%= itemid %>" size="11" maxlength="11" onKeyPress="if (event.keyCode == 13) SubmitSearch();">
		</td>
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="�˻�" onClick="javascript:SubmitSearch();">
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">
			CS���� :
			<select class="select" name="searchtype">
				<option value="">��ü</option>
				<option value="jupsuChulgo"   <% if (searchtype = "jupsuChulgo") then %>selected<% end if %>>��ȯ CS����</option>
				<option value="confirmChulgo" <% if (searchtype = "confirmChulgo") then %>selected<% end if %>>��ȯ ��üȮ��</option>
				<option value="jupsuReturn" <% if (searchtype = "jupsuReturn") then %>selected<% end if %>>��ǰ����</option>
			</select>
			&nbsp;
			<input type=checkbox name="showimage" value="Y" <% if (showimage = "Y") then %>checked<% end if %>> �̹��� ǥ��
		</td>
	</tr>
	</form>
</table>

<p>
<table width="100%" border="0" class="a" cellpadding="3" cellspacing="1" >
<tr>
    <td>
    * CS���� �̻� ��üó���Ϸ� ���� ���±����� ��ǰ���� ǥ���մϴ�.(�ֱ� 3����)
    </td>
    <td align="right" width="100"><img src="/images/btn_excel.gif" width=75 onClick="xlDown();" style="cursor:pointer"></td>
</tr>
</table>
<p>

	<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
	    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
			<td width="60">��ǰ�ڵ�</td>
			<% if (showimage = "Y") then %>
				<td width="50">�̹���</td>
			<% end if %>
			<td>��ǰ��</td>
			<td width="140">�ɼǸ�</td>
			<td width="85">��ȯ CS����</td>
			<td width="85">��ȯ ��üȮ��</td>
			<td width="85">��ǰ����</td>
	    </tr>
<% if oitem.FresultCount<1 then %>
	    <tr bgcolor="#FFFFFF">
	    	<td colspan="6" align="center">[�˻������ �����ϴ�.]</td>
	    </tr>
<% end if %>
<% if oitem.FresultCount > 0 then %>
	<%
	jupsuChulgoSUM = 0
	confirmChulgoSUM = 0
	jupsuReturnSUM = 0
	%>
    <% for i=0 to oitem.FresultCount-1 %>
    	<% if (oitem.FItemList(i).Fisusing = "N") then %>
    	<tr class="a" height="25" bgcolor="<%= adminColor("gray") %>">
		<% else %>
		<tr class="a" height="25" bgcolor="#FFFFFF">
		<% end if %>
			<td align="center"><%= oitem.FItemList(i).Fitemid %></td>
			<% if (showimage = "Y") then %>
				<td align="center">
					<img src="<%= oitem.FItemList(i).FImgSmall %>" width="50" height="50" border="0" alt="">
				</td>
			<% end if %>
			<td align="left">
				<a href="http://www.10x10.co.kr/shopping/category_prd.asp?itemid=<%= oitem.FItemList(i).Fitemid %>" target="_blank"><% =oitem.FItemList(i).Fitemname %>&nbsp;&nbsp;</a>
			</td>
			<td align="center">
				<%= oitem.FItemList(i).Fitemoptionname %>
			</td>
		    <td align="center">
				<%= FormatNumber(oitem.FItemList(i).FjupsuChulgo,0) %>
		    </td>
		    <td align="center">
				<%= FormatNumber(oitem.FItemList(i).FconfirmChulgo,0) %>
		    </td>
		    <td align="center">
				<%= FormatNumber(oitem.FItemList(i).FjupsuReturn,0) %>
		    </td>
		</tr>
			<%
			jupsuChulgoSUM = jupsuChulgoSUM + oitem.FItemList(i).FjupsuChulgo
			confirmChulgoSUM = confirmChulgoSUM + oitem.FItemList(i).FconfirmChulgo
			jupsuReturnSUM = jupsuReturnSUM + oitem.FItemList(i).FjupsuReturn
			%>
		<% next %>
		<tr class="a" height="40" bgcolor="#FFFFFF">
			<td align="center" colspan="<% if (showimage = "Y") then %>4<% else %>3<% end if %>"></td>
		    <td align="center">
				<b><%= FormatNumber(jupsuChulgoSUM,0) %></b>
		    </td>
		    <td align="center">
				<b><%= FormatNumber(confirmChulgoSUM,0) %></b>
		    </td>
		    <td align="center">
				<b><%= FormatNumber(jupsuReturnSUM,0) %></b>
		    </td>
		</tr>
	</table>
<% end if %>

<!-- ǥ �ϴܹ� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
    <tr valign="bottom" height="10">
        <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_08.gif"></td>
        <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
    </tr>
</table>
<!-- ǥ �ϴܹ� ��-->
<iframe name="iiframeXL" name="iiframeXL" width="0" height="0" frameborder=0 scrolling=no marginheight=0 marginwidth=0 align=center></iframe>
<form name=xlfrm method=post action="">
<input type="hidden" name="searchtype" value="<%= searchtype %>">
<input type="hidden" name="itemid" value="<%= itemid %>">
</form>
<!-- #include virtual="/designer/lib/designerbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
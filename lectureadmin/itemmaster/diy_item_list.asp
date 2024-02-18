<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/designer/lib/designerbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/academy/lib/classes/DIYShopItem/DIYitemCls.asp"-->
<%

dim itemid, makerid, itemname
dim sellyn, isusing, limityn, mwdiv
dim page

itemid  = RequestCheckVar(request("itemid"),10)
makerid = RequestCheckVar(request("makerid"),32)
itemname = RequestCheckVar(request("itemname"),32)

sellyn  = RequestCheckVar(request("sellyn"),10)
isusing = RequestCheckVar(request("isusing"),10)
limityn = RequestCheckVar(request("limityn"),10)
mwdiv = RequestCheckVar(request("mwdiv"),10)

page = RequestCheckVar(request("page"),10)



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
oitem.FRectLimityn = limityn
oitem.FRectMWDiv = mwdiv
oitem.FPageSize = 30
oitem.FCurrPage = page


if (sellyn <> "A") then
    oitem.FRectSellYN = sellyn
end if

if (isusing <> "A") then
    oitem.FRectIsUsing = isusing
end if


oitem.GetItemList

dim i

%>
<script>
function viewBySite(itemid){
    <% if (now()<#2016-09-06#) then %>
   // alert('9�� 5�� ����Ʈ�� ����˴ϴ�. ���� ���̴� �������� 9��5�� ���� �������� �ٸ��� �̹��� ������ 9�� 5�� ���� �Ͻñ� �ٶ��ϴ�.\r\n\r\n���������� �̹���Ÿ�� ���簢��, \r\n����� ������ ��ǰ�̹���Ÿ�� ���簢��');
    <% end if %>
    window.open('<%=wwwFingers%>/diyshop/shop_prd.asp?itemid='+itemid,'_blank'); 
}

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


// ============================================================================
// �⺻��������
function editItemInfo(itemid) {
	var param = "itemid=" + itemid + "&fingerson=on";

	popwin = window.open('diy_item_infomodify.asp?' + param ,'editItemInfo','width=900,height=600,scrollbars=yes,resizable=yes');
	popwin.focus();
}

// ============================================================================
// �ɼǼ���
function editItemOption(itemid) {
	var param = "itemid=" + itemid;

	popwin = window.open('diy_item_optionmodify.asp?' + param ,'editItemOption','width=900,height=600,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function editSimpleItemOption(itemid) {
	var param = "itemid=" + itemid;

	popwin = window.open('/academy/comm/pop_diy_simpleitemedit.asp?' + param ,'editSimpleItemOption','width=500,height=650,scrollbars=yes,resizable=yes');
	popwin.focus();
}

// ============================================================================
// �̹�������
function editItemImage(itemid) {
	var param = "itemid=" + itemid;

	popwin = window.open('diy_item_imagemodify.asp?' + param ,'editItemImage','width=900,height=600,scrollbars=yes,resizable=yes');
	popwin.focus();
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
			&nbsp;
			��ǰ�� :
			<input type="text" class="text" name="itemname" value="<%= itemname %>" size="20" onKeyPress="if (event.keyCode == 13) SubmitSearch();"><br>
		</td>
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="�˻�" onClick="javascript:SubmitSearch();">
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">
			�Ǹ�:<% drawSelectBoxSellYN "sellyn", sellyn %>
			&nbsp;
	     	����:<% drawSelectBoxLimitYN "limityn", limityn %>
	     	&nbsp;
	     	�ŷ�����:<% drawSelectBoxMWU "mwdiv", mwdiv %>
	     	&nbsp;
		</td>
	</tr>
	</form>
</table>

<p>

	<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
	    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
			<td width="60">��ǰ�ڵ�</td>
			<td width="50">�̹���</td>
			<td>��ǰ��</td>
			<td width="30">�ŷ�<br>����</td>
			<td width="30">�Ǹ�<br>����</td>
			<td width="40">����<br>����</td>
			<td width="60">�ǸŰ�</td>
			<td width="60">���ް�</td>
			<td width="50">�⺻<br>����</td>
			<td width="50">�̹���</td>
			<td width="70">�����Ǹ�<br>�Ǹſ���</td>
	    </tr>
<% if oitem.FresultCount<1 then %>
	    <tr bgcolor="#FFFFFF">
	    	<td colspan="13" align="center">[�˻������ �����ϴ�.]</td>
	    </tr>
<% end if %>
<% if oitem.FresultCount > 0 then %>
    <% for i=0 to oitem.FresultCount-1 %>
    	<% if (oitem.FItemList(i).Fisusing = "N") then %>
    	<tr class="a" height="25" bgcolor="<%= adminColor("gray") %>">
		<% else %>
		<tr class="a" height="25" bgcolor="#FFFFFF">
		<% end if %>
			<td align="center"><%= oitem.FItemList(i).Fitemid %></td>
			<td align="center"><img src="<%= oitem.FItemList(i).Fsmallimage %>" width="50" height="50" border="0" alt=""></td>
			<% if (FALSE) then %>
			<td align="left"><% =oitem.FItemList(i).Fitemname %>&nbsp;&nbsp;<a href="<%=wwwFingers%>/diyshop/shop_prd.asp?itemid=<%= oitem.FItemList(i).Fitemid %>" target="_blank"><font color="blue">(Ȯ���ϱ�)</font></a></td>
		    <% else %>
		    <td align="left"><% =oitem.FItemList(i).Fitemname %>&nbsp;&nbsp;<a href="javascript:viewBySite('<%= oitem.FItemList(i).Fitemid %>');" ><font color="blue">(Ȯ���ϱ�)</font></a></td>
	        <% end if %>
			<td align="center">
				<font color="<%= mwdivColor(oitem.FItemList(i).Fmwdiv) %>"><%= mwdivName(oitem.FItemList(i).Fmwdiv) %></font>
			</td>

			<td align="center"><%= fnColor(oitem.FItemList(i).Fsellyn,"yn") %></td>
			<td align="center">
        		<% if (oitem.FItemList(i).Flimityn = "Y") then %>
             		<%= fnColor(oitem.FItemList(i).Flimityn,"yn") %>
             		<br>(<%= (oitem.FItemList(i).Flimitno - oitem.FItemList(i).Flimitsold) %>)
        		<% else %>
              		<%= fnColor(oitem.FItemList(i).Flimityn,"yn") %>
       			<% end if %>
			</td>
			<td align="right"><%= FormatNumber(oitem.FItemList(i).Fsellcash,0) %></td>
			<td align="right"><%= FormatNumber(oitem.FItemList(i).Fbuycash,0) %></td>
		    <td align="center">
		    	<a href="javascript:editItemInfo('<%= oitem.FItemList(i).FItemId %>')">
		    	<img src="/images/icon_modify.gif" border="0" align="absbottom">
		    	</a>
		    </td>
		    <td align="center">
		    	<a href="javascript:editItemImage('<%= oitem.FItemList(i).FItemId %>')">
		    	<img src="/images/icon_modify.gif" border="0" align="absbottom">
		    	</a>
		    </td>
		    <td align="center">
        <% if (oitem.FItemList(i).Fmwdiv = "U") then %>
		      	<a href="javascript:editSimpleItemOption('<%= oitem.FItemList(i).FItemId %>')">
		      	<img src="/images/icon_modify.gif" border="0" align="absbottom">
		      	</a>
        <% else %>
		      	<a href="javascript:editSimpleItemOption('<%= oitem.FItemList(i).FItemId %>')">
		      	<b>[</b>������û<b>]</b>
		      	</a>
        <% end if %>

		    </td>
		</tr>
		<% next %>
	</table>
<% end if %>

<!-- ǥ �ϴܹ� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
    <tr valign="top" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="center">
            <% if oitem.HasPreScroll then %>
			<a href="javascript:NextPage('<%= oitem.StartScrollPage-1 %>')">[pre]</a>
    		<% else %>
    			[pre]
    		<% end if %>

    		<% for i=0 + oitem.StartScrollPage to oitem.FScrollCount + oitem.StartScrollPage - 1 %>
    			<% if i>oitem.FTotalpage then Exit for %>
    			<% if CStr(page)=CStr(i) then %>
    			<font color="red">[<%= i %>]</font>
    			<% else %>
    			<a href="javascript:NextPage('<%= i %>')">[<%= i %>]</a>
    			<% end if %>
    		<% next %>

    		<% if oitem.HasNextScroll then %>
    			<a href="javascript:NextPage('<%= i %>')">[next]</a>
    		<% else %>
    			[next]
    		<% end if %>
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    <tr valign="bottom" height="10">
        <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_08.gif"></td>
        <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
    </tr>
</table>
<!-- ǥ �ϴܹ� ��-->
<!-- #include virtual="/designer/lib/designerbodytail.asp"-->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<% option Explicit %>
<!-- #include virtual="/common/incSessionBctId.asp" -->
<!-- #include virtual="/common/lib/popheader.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/classes/items/basicItemInfocls.asp" -->
<%
'// ���� ���� //
dim makerid, itemid, gubun, SQL, addSQL
dim page, oItem, ix
dim regstate, research

'// �Ķ���� ���� //
makerid = requestCheckVar(request("makerid"),32)
itemid  = getNumeric(requestCheckVar(Request("itemid"),9))
page    = getNumeric(requestCheckVar(Request("page"),9))
regstate = requestCheckVar(Request("regstate"),9)
research = requestCheckVar(Request("research"),2)
dim IsACADEMYDIY : IsACADEMYDIY  = (requestCheckVar(Request("tp"),16)="academydiy")

if page="" then page=1 else page=Cint(page)
if ((research="") and (regstate="")) then regstate="F"

// ��ü�ΰ�� ��ü ��ǰ�� ����.
if (C_IS_Maker_Upche) then
	makerid = session("ssBctId")
end if

'��ǰ�ڵ� ��ȿ�� �˻�(2008.07.15;������)
if itemid<>"" then
	if Not(isNumeric(itemid)) then
		Response.Write "<script language=javascript>alert('[" & itemid & "]��(��) ��ȿ�� ��ǰ�ڵ尡 �ƴմϴ�.');history.back();</script>"
		dbget.close()	:	response.End
	end if
end if

'// ���� ��ǰ ��� ���� //
set oItem = new CItemlist

oitem.FPageSize = 10
oitem.FCurrPage = page
oitem.FRectRegState = regstate

if makerid<>"" then
	oitem.FRectMakerId = makerid
end if
if itemid<>"" then
	oitem.FRectItemId = itemid
end if
oitem.ProductList
%>
<script language='javascript'>
<!--
	// ���� ����
	function inputItemCont(icd)
	{
	    <% if (IsACADEMYDIY) then %>
		self.location = "basic_item_info_list_insert_academydiy.asp?itemId=" + icd + "&regstate=<%= regstate %>";
		<% else %>
		self.location = "basic_item_info_list_insert.asp?itemId=" + icd + "&regstate=<%= regstate %>";
		<% end if %>
	}

	// ������ �̵�
	function NextPage(ipage)
	{
		document.frm.page.value= ipage;
		document.frm.submit();
	}

	// �˻�!
	function search()
	{
		document.frm.page.value= "1";
		document.frm.submit();
	}
//-->
</script>

<!-- ǥ ��ܹ� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
   	<form name="frm" method="get" action="pop_basic_item_info_list.asp">
	<input type="hidden" name="page" value="">
	<input type="hidden" name="research" value="on">
	<input type="hidden" name="tp" value="<%= Request("tp") %>">
   	<tr height="10" valign="bottom">
	        <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
	        <td background="/images/tbl_blue_round_02.gif"></td>
	        <td background="/images/tbl_blue_round_02.gif"></td>
	        <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<tr height="25" valign="bottom">
	        <td background="/images/tbl_blue_round_04.gif"></td>
	        <td>
	        	<input type="radio" name="regstate" value="F" <% if regstate="F" then response.write "checked" %> >��ϿϷ� �� ��ǰ
				<input type="radio" name="regstate" value="W" <% if regstate="W" then response.write "checked" %> >��ϴ�� �� ��ǰ
				<br>
				��ǰ�ڵ� <input type="text" name="itemid" value="<%=itemid%>" size="5">
				<%
					Select Case session("ssBctDiv")
						Case "9999"
							Response.Write "�귣�� : <b>" & session("ssBctCname") & "</b>"
						Case Else
							Response.Write "�귣�� : "
							Call drawSelectBoxDesignerwithName("makerid", makerid)
					end Select
				%>
	        </td>
	        <td align="right">
	        	<a href="javascript:search();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
	        </td>
	        <td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	</form>
</table>
<!-- ǥ ��ܹ� ��-->


<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td>�귣��ID</td>
		<td width="50">�̹���</td>
		<td width="50">��ǰ�ڵ�</td>
		<td>��ǰ��</td>
		<td>�ǸŰ�</td>
		<td>�Ǹſ���</td>
		<td>��뿩��</td>
		<td>����</td>
	</tr>
	<% if oitem.FresultCount<1 then %>
	<tr bgcolor="#FFFFFF">
		<td colspan="8" align="center">
			<br>[�˻������ �����ϴ�.]<br><br>
			<span onClick="self.close()" style="cursor:pointer">[�ݱ�]</span>
		</td>
	</tr>
	<% else %>
		
	<% for ix=0 to oitem.FresultCount-1 %>
	<tr align="center" bgcolor="#FFFFFF">
		<td><%= oitem.FItemList(ix).Fmakerid %></td>
		<td><img src="<%= oitem.FItemList(ix).FImgSmall %>" width="50" height="50" border="0"</td>
		<td><%= oitem.FItemList(ix).Fitemid %></td>
		<td><% = oitem.FItemList(ix).Fitemname %></td>
		<td align=right><%= FormatNumber(oitem.FItemList(ix).Fsellcash,0) %></td>
		<td><%= oitem.FItemList(ix).Fsellyn %></td>
		<td><%= oitem.FItemList(ix).Fisusing %></td>
		<td><a href="javascript:inputItemCont('<%= oitem.FItemList(ix).Fitemid %>')"><img src="/images/icon_use.gif" border="0" align="absbottom"></a></td>
	</tr>
	<% next %>
	
	<% end if %>
</table>


<!-- ǥ �ϴܹ� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
    <tr valign="top" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="center">
        	<%
				'@ ���� ������ ���
				if oitem.HasPreScroll then
					Response.Write "<a href=""javascript:NextPage('" & oitem.StarScrollPage-1 & "')"">[pre]</a>"
				else
					Response.Write "[pre]"
				end if
	
				'@ ������ ��ȣ ���
				for ix=(0 + oitem.StarScrollPage) to (oitem.StarScrollPage + oitem.FScrollCount - 1)
	
					if (ix > oitem.FTotalpage) then Exit for
					if CStr(ix) = CStr(oitem.FCurrPage) then
						Response.Write "<font color='red'>[" & ix & "]</font>"
					else
						Response.Write "<a href=""javascript:NextPage('" & ix & "')"">[" & ix & "]</a>"
					end if
				next
	
				'@ ���� ������ ���
				if oitem.HasNextScroll then
					Response.Write "<a href=""javascript:NextPage('" & ix & "')"">[next]</a>"
				else
					Response.Write "[next]"
				end if
			%>
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
	
<%
set oitem = Nothing
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/common/lib/poptail.asp"-->
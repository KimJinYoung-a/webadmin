<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/classes/items/extsiteitemcls.asp"-->

<%
dim notmatch, research, page, cdl
notmatch = request("notmatch")
research = request("research")
page     = request("page")
cdl      = RequestCheckVar(request("cdl"),3)

if ((research="") and (notmatch="")) then notmatch="on"
if (page="") then page=1

dim oInterParkitem
set oInterParkitem = new CExtSiteItem
oInterParkitem.FRectNotMatchCategory = notmatch
oInterParkitem.FRectCate_large = cdl

'if (cdl<>"") then
    oInterParkitem.GetInterParkCategoryMachingList
'end if

dim i
%>
<script language='javascript'>
function MatcheDispCate(cdl,cdm,cdn){
    var popwin = window.open('InterParkMatcheDispCate.asp?cdl=' + cdl + '&cdm=' + cdm +'&cdn=' + cdn,'MatcheDispCate','width=800,height=400,scrollbars=yes,resizable=yes');
    popwin.focus();
}

function changecontent(){
    //nothing
}

function popInterparkCate()
{
	window.open('Pop_InterPark_Category.asp','interparkcate','width=900,height=527,scrollbars=yes');
}
</script>
<table width="100%" border="0" cellpadding="5" cellspacing="1" bgcolor="#EEEEEE">
	<form name="frm" method="get" action="">
	<input type="hidden" name="page" value="1">
	<input type="hidden" name="research" value="on">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<tr >
		<td class="a">
    		<input type="checkbox" name="notmatch" <%= ChkIIF(notmatch="on","checked","") %> >��Ī �ȵ� ���� �� ������� ī�װ� ��Ī��
    		&nbsp;
    		ī�װ� : <% call DrawSelectBoxCategoryLarge("cdl",cdl) %>
		</td>
		<td class="a" align="right">
			<a href="javascript:document.frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
		</td>
	</tr>
	</form>
</table>

<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="#CCCCCC">
<tr bgcolor="#FFFFFF">
    <tD colspan="9">
        <table width="600" class="a" cellpadding="2" cellspacing="1"bgcolor="#CCCCCC">
        <tr bgcolor="#FFFFFF"><td width="100" bgcolor='#FFCCCC'></td><td>����ī�װ�</td><td bgcolor='#CCCCCC' width="100" ></td><td>���þ���ī�װ�</td><td bgcolor='#CCCCFF' width="100" ></td><td>��Ī�ȵ�ī�װ�</td></tr>
        </table>    
    </td>
</tr>
<tr bgcolor="#FFFFFF">
	<td colspan="9">
		<table width="100%" cellpadding="0" cellspacing="0" border="0" class="a">
		<tr>
			<td><input type="button" class="button" value="InterPark ī�װ� ���� �� �߰�" onClick="popInterparkCate()"></td>
			<td align="right" height="30">page: <%= FormatNumber(page,0) %> / <%= FormatNumber(oInterParkitem.FTotalPage,0) %> �ѰǼ�: <%= FormatNumber(oInterParkitem.FTotalCount,0) %></td>
		</tr>
		</table>
	</td>
</tr>
<tr align="center" bgcolor="#F3F3FF" height="20">
	<td width="100">Ten ī���ڵ�</td>
	<td width="100">��з�</td>
	<td width="100">�ߺз�</td>
	<td width="100">�Һз�</td>
	<td width="100">��ǰ��</td>
	<td width="100">���ް���ڵ�</td>
	<td width="100">iPark ����1</td>
	<td width="100">iPark �귣������1</td>
	<td width="100">iPark ����1(�ѱ�)</td>
</tr>
<% for i=0 to oInterParkitem.FResultCount-1 %>
<tr align="center" bgcolor="#FFFFFF">
    <td><%= oInterParkitem.FItemList(i).FCate_Large %><%= oInterParkitem.FItemList(i).FCate_Mid %><%= oInterParkitem.FItemList(i).FCate_Small %></td>
    <td><%= oInterParkitem.FItemList(i).Fnmlarge %></td>
    <td><%= oInterParkitem.FItemList(i).FnmMid %></td>
    <td><%= oInterParkitem.FItemList(i).FnmSmall %></td>
    <td><%= oInterParkitem.FItemList(i).FItemCnt %></td>
    <td><%= oInterParkitem.FItemList(i).getSupplyCtrtSeqName %></td>
    <td <% if oInterParkitem.FItemList(i).FIparkCateDispyn="N" then 
            response.write "bgcolor='#CCCCCC'" 
           elseif oInterParkitem.FItemList(i).FIparkCateDispyn="D" then 
            response.write "bgcolor='#FFCCCC'"
           elseif IsNULL(oInterParkitem.FItemList(i).FIparkCateDispyn) then 
            response.write "bgcolor='#CCCCFF'"
           end if
        %> >
        <% if oInterParkitem.FItemList(i).IsNotMatchedDispcategory then %>
        <input type="button" class="button" value="���" onclick="MatcheDispCate('<%= oInterParkitem.FItemList(i).FCate_Large %>','<%= oInterParkitem.FItemList(i).FCate_Mid %>','<%= oInterParkitem.FItemList(i).FCate_Small %>');">
        <% else %>
        <a href="javascript:MatcheDispCate('<%= oInterParkitem.FItemList(i).FCate_Large %>','<%= oInterParkitem.FItemList(i).FCate_Mid %>','<%= oInterParkitem.FItemList(i).FCate_Small %>');"><%= oInterParkitem.FItemList(i).Finterparkdispcategory %></a>
        <% end if %>
    </td>
    <td>
        <% if oInterParkitem.FItemList(i).IsNotMatchedStorecategory then %>
        <input type="button" class="button" value="���" onclick="MatcheDispCate('<%= oInterParkitem.FItemList(i).FCate_Large %>','<%= oInterParkitem.FItemList(i).FCate_Mid %>','<%= oInterParkitem.FItemList(i).FCate_Small %>');">
        <% else %>
        <a href="javascript:MatcheDispCate('<%= oInterParkitem.FItemList(i).FCate_Large %>','<%= oInterParkitem.FItemList(i).FCate_Mid %>','<%= oInterParkitem.FItemList(i).FCate_Small %>');"><%= oInterParkitem.FItemList(i).Finterparkstorecategory %></a>
        <% end if %>
    </td>
    <td><%= oInterParkitem.FItemList(i).FinterparkdispcategoryText %></td>
</tr> 
<% next %>
</table>

<%
set oInterParkitem = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
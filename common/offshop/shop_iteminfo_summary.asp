<%@ language=vbscript %>
<% option explicit %>
<%
	Response.AddHeader "Cache-Control","no-cache"
	Response.AddHeader "Expires","0"
	Response.AddHeader "Pragma","no-cache"
%>
<%
'###########################################################
' Description : ��ǰ�ֹ��˻�
' Hieditor : 2011.08.22 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/common/lib/commonbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshopitemcls.asp"-->
<!-- #include virtual="/lib/classes/offshop/stock/shortagestock_cls.asp" -->
<!-- #include virtual="/lib/classes/items/adminshoppingbag/adminshoppingbag_cls.asp" -->
<%
dim page , shopid , isusing , makerid , itemid , itemname , generalbarcode , i , sell7days ,shopsuplycash ,buycash
dim cdl , cdm , cds , shortagetype , comm_cd ,includepreorder ,research , parameter , ipgo , order
dim pagesize
	page = requestCheckVar(request("page"),10)
	pagesize = requestCheckVar(request("pagesize"),10)
    research = requestCheckVar(request("research"),2)
    isusing = requestCheckVar(request("isusing"),1)
    makerid = requestCheckVar(request("makerid"),32)
    itemid = requestCheckVar(request("itemid"),10)
    itemname = requestCheckVar(request("itemname"),124)
    generalbarcode = requestCheckVar(request("generalbarcode"),32)
    comm_cd = requestCheckVar(request("comm_cd"),32)
    cdl = requestCheckVar(request("cdl"),3)
    cdm = requestCheckVar(request("cdm"),3)
    cds = requestCheckVar(request("cds"),3)
    shortagetype = requestCheckVar(request("shortagetype"),32)
    includepreorder = requestCheckVar(request("includepreorder"),32)
    sell7days = requestCheckVar(request("sell7days"),32)
    ipgo = requestCheckVar(request("ipgo"),32)
	shopid = requestCheckVar(request("shopid"),32)
    order = requestCheckVar(request("order"),32)

if page="" then page=1
if pagesize="" then pagesize=1000
if pagesize > 1000 then pagesize = 1000
isusing = "Y"

'/�����ϰ�� ���� ���常 ��밡��
if (C_IS_SHOP) then
	'/���α��� ���� �̸�
	'if getlevel_sn("",session("ssBctId")) > 6 then
		shopid = C_STREETSHOPID
	'end if
else
	if (C_IS_Maker_Upche) then
		shopid = session("ssBctID")
	else
		if (Not C_ADMIN_USER) then

		else

		end if
	end if
end if

'if shopid = "" then shopid = "streetshop011"

parameter = "page="&page&"&research="&research&"&shopid="&shopid&"&isusing="&isusing&"&makerid="&makerid&"&itemid="&itemid&"&itemname="&itemname&"&sell7days="&sell7days&""
parameter = parameter & "&generalbarcode="&generalbarcode&"&comm_cd="&comm_cd&"&cdl="&cdl&"&cdm="&cdm&"&cds="&cds&"&shortagetype="&shortagetype&"&includepreorder="&includepreorder
parameter = parameter & "&ipgo="&ipgo&"&order="&order&""

dim oshortage
set oshortage  = new cshortagestock_list
    oshortage.FPageSize = pagesize
    oshortage.FCurrPage = page
    oshortage.Frectshopid = shopid

    if shopid <> "" then
		if (LCASE(shopid) <> "wholesale1043") and (LCASE(shopid) <> "wholesaletest") then
			response.write "<script language='javascript'>"
			response.write "    alert('�߸��� �����Դϴ�.');"
			response.write "</script>"

			db3_dbget.close:dbget.Close:response.end
		end if

		oshortage.fnewitemstock_list_datamart2
    else
        response.write "<script language='javascript'>"
        response.write "    alert('������ ������ �ּ���');"
        response.write "</script>"
    end if

dim IsUpcheWitakItem
if (makerid<>"") and (shopid<>"") then
    IsUpcheWitakItem = (GetShopBrandContract(shopid,makerid)="B012")
end if

Dim isellyn, ilimitNo
%>

<script type='text/javascript'>

//�˻���ư
function reg(page){
    frm.page.value=page;
    frm.submit();
}

</script>

<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="research" value="on">
<input type="hidden" name="page" value="<%= page %>">
<tr align="center" bgcolor="#FFFFFF" >
    <td rowspan="1" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
    <td align="left">
        ���� : <%= shopid %>
        &nbsp;
        ��뿩��: Y
		&nbsp;
		������������:
		<select class="select" name="pagesize">
			<option value="50" <% if (pagesize = "50") then %>selected<% end if %> >50</option>
			<option value="100" <% if (pagesize = "100") then %>selected<% end if %> >100</option>
			<option value="500" <% if (pagesize = "500") then %>selected<% end if %> >500</option>
			<option value="1000" <% if (pagesize = "1000") then %>selected<% end if %> >1000</option>
		</select>
    </td>
    <td rowspan="1" width="50" bgcolor="<%= adminColor("gray") %>">
        <input type="button" class="button_s" value="�˻�" onClick="javascript:reg('');">
    </td>
</tr>
</form>
</table>
<!-- �˻� �� -->
<br>

<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
<tr>
    <td align="left">
		* ���� ���� 01�� 50�п� ������Ʈ �˴ϴ�.<br>
		* �Ǹſ��� : Y/N/S (S = �Ͻ�ǰ��)<br>
		* ��۱��� : T/U (T = �ٹ����ٹ��, U = ��ü���)
    </td>
    <td align="right">

    </td>
</tr>
</table>
<!-- �׼� �� -->

<p>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
    <td colspan="25">
        �˻���� : <b><%= oshortage.FTotalcount %></b>
        &nbsp;
        <b><%= page %> / <%= oshortage.FTotalpage %></b>
    </td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    <td>
    	����ó
    </td>
    <td>�귣��</td>
    <td>��ǰ�ڵ�</td>
    <td>�̹���</td>
    <td>��ǰ��<br><font color="blue">[�ɼǸ�]</font></td>
	<% if not(C_IS_Maker_Upche) and not(C_IS_SHOP) then %>
    	<td>���԰�</td>
    <% end if %>
    <td>�ǸŰ�</td>
    <td>����<br>���ް�</td>
    <td>����<br>����</td>
    <td>��뿩��</td>
    <td>�Ǹſ���</td>

	<td>��۱���</td>
	<td>����</td>
	<td>������</td>
	<td>����</td>
	<td>������</td>

	<td>��������</td>
	<td>��������</td>

    <td>���</td>
</tr>
<% if oshortage.FresultCount > 0 then %>
<%
for i=0 to oshortage.FresultCount -1

shopsuplycash = oshortage.FItemList(i).GetFranchiseSuplycash
buycash		  = oshortage.FItemList(i).GetFranchiseBuycash

isellyn = oshortage.FItemlist(i).Fsellyn
ilimitNo = 0

if oshortage.FItemList(i).Foptlimityn="Y" then
    ilimitNo = oshortage.FItemlist(i).Foptlimitno-oshortage.FItemlist(i).Foptlimitsold-5
    if ilimitNo<1 then ilimitNo=0
end if

if (isellyn="Y") and (oshortage.FItemList(i).Foptlimityn="Y") and (ilimitNo<1) then
    isellyn="S"
end if

if (oshortage.FItemlist(i).Foptisusing="N" or oshortage.FItemlist(i).Foptsellyn="N") then
    isellyn="N"
end if
%>
<form method="get" action="" name="frmBuyPrc<%=i%>">

<% if oshortage.FItemlist(i).Fisusing="N" then %>
<tr bgcolor="#EEEEEE" align="center">
<% else %>
<tr bgcolor="#FFFFFF" align="center">
<% end if %>
<input type="hidden" name="comm_cd" value="<%= oshortage.FItemlist(i).fcomm_cd %>">
<input type="hidden" name="itemgubun" value="<%= oshortage.FItemlist(i).fitemgubun %>">
<input type="hidden" name="itemid" value="<%= oshortage.FItemlist(i).fitemid %>">
<input type="hidden" name="itemoption" value="<%= oshortage.FItemlist(i).fitemoption %>">
<input type="hidden" name="shopitemprice" value="<%= shopsuplycash %>">
<input type="hidden" name="shopbuyprice" value="<%= buycash %>">
<input type="hidden" name="itemname" value="<%= oshortage.FItemlist(i).fshopitemname %>">
<input type="hidden" name="itemoptionname" value="<%= oshortage.FItemlist(i).fshopitemoptionname %>">
<input type="hidden" name="makerid" value="<%= oshortage.FItemlist(i).fmakerid %>">
<input type="hidden" name="shopid" value="<%= oshortage.FItemlist(i).fshopid %>">
<input type="hidden" name="sellyn" value="<%= isellyn %>">
<input type="hidden" name="deliverytype" value="<%= oshortage.FItemlist(i).Fdeliverytype %>">
<input type="hidden" name="itemweight" value="<%= oshortage.FItemlist(i).FitemWeight %>">
<input type="hidden" name="sourcearea" value="<%= oshortage.FItemlist(i).Fsourcearea %>">
<input type="hidden" name="itemsource" value="<%= oshortage.FItemlist(i).Fitemsource %>">
<input type="hidden" name="itemsize" value="<%= oshortage.FItemlist(i).Fitemsize %>">
<input type="hidden" name="limityn" value="<%= oshortage.FItemlist(i).Foptlimityn %>">
<input type="hidden" name="limitno" value="<%= ilimitNo %>">
<input type="hidden" name="mmgcate" value="<%= oshortage.FItemList(i).fcatecdl %><%= oshortage.FItemList(i).fcatecdm %><%= oshortage.FItemList(i).fcatecdn %>">
<input type="hidden" name="dispcate" value="<%= oshortage.FItemList(i).Fdispcatecode %>">
<input type="hidden" name="mimgurl" value="<%= oshortage.FItemList(i).GetBasicImage %>">
    <td>
        <%= GetdeliverGubunName(oshortage.FItemlist(i).fcomm_cd) %><br>(<%= GetJungsanGubunName(oshortage.FItemlist(i).fcomm_cd) %>)
    </td>
    <td>
        <%= oshortage.FItemlist(i).fmakerid %>
    </td>
    <td>
        <%= oshortage.FItemlist(i).Fitemgubun %><%= CHKIIF(oshortage.FItemlist(i).Fitemid>=1000000,Format00(8,oshortage.FItemlist(i).Fitemid),Format00(6,oshortage.FItemlist(i).Fitemid)) %><%= oshortage.FItemlist(i).Fitemoption %>
    </td>
    <td><img src="<%= oshortage.FItemlist(i).GetImageSmall %>" width=50 height=50 border=0></td>
    <td>
        <%= oshortage.FItemlist(i).fshopitemname %><Br>
        <% if oshortage.FItemlist(i).fshopitemoptionname <> "" then %>
            <font color="blue">[<%=oshortage.FItemlist(i).fshopitemoptionname%>]<font>
        <% end if %>
    </td>
	<% if not(C_IS_Maker_Upche) and not(C_IS_SHOP) then %>
	    <td>
	        <%= FormatNumber(oshortage.FItemlist(i).fshopsuplycash,0) %>
	    </td>
	<% end if %>
    <td>
        <%= FormatNumber(oshortage.FItemlist(i).fshopitemprice,0) %>
    </td>
    <td>
        <%= FormatNumber(oshortage.FItemlist(i).fshopbuyprice,0) %>
    </td>
    <td>
		<% if oshortage.FItemList(i).Fshopitemprice<>0 then %>
		<%= 100-(CLng(shopsuplycash/oshortage.FItemList(i).Fshopitemprice*10000)/100) %> %
		<% end if %>
    </td>
    <td><%= oshortage.FItemlist(i).Fisusing %></td>
    <td><%= oshortage.FItemlist(i).Fsellyn %></td>
	<td><%= oshortage.FItemlist(i).Fdeliverytype %></td>
	<td><%= oshortage.FItemlist(i).FitemWeight %></td>
	<td><%= oshortage.FItemlist(i).Fsourcearea %></td>
	<td><%= oshortage.FItemlist(i).Fitemsource %></td>
	<td><%= oshortage.FItemlist(i).Fitemsize %></td>
	<td><%= oshortage.FItemlist(i).Foptlimityn %></td>
	<td><%= ilimitNo %></td>

    <td></td>
</tr>
</form>
<% next %>

<tr height="25" bgcolor="FFFFFF">
    <td colspan="25" align="center">
        <% if oshortage.HasPreScroll then %>
            <span class="list_link"><a href="javascript:reg(<%=oshortage.StartScrollPage-1%>)">[pre]</a></span>
        <% else %>
        [pre]
        <% end if %>
        <% for i = 0 + oshortage.StartScrollPage to oshortage.StartScrollPage + oshortage.FScrollCount - 1 %>
            <% if (i > oshortage.FTotalpage) then Exit for %>
            <% if CStr(i) = CStr(oshortage.FCurrPage) then %>
            <span class="page_link"><font color="red"><b><%= i %></b></font></span>
            <% else %>
            <a href="javascript:reg(<%=i%>);" class="list_link"><font color="#000000"><%= i %></font></a>
            <% end if %>
        <% next %>
        <% if oshortage.HasNextScroll then %>
            <span class="list_link"><a href="javascript:reg(<%=i%>);">[next]</a></span>
        <% else %>
        [next]
        <% end if %>
    </td>
</tr>

<% else %>

<tr bgcolor="#FFFFFF">
    <td colspan="25" align="center">[�˻������ �����ϴ�.]</td>
</tr>
<% end if %>
</table>
<%
    set oshortage = nothing
%>
<!-- #include virtual="/common/lib/commonbodytail.asp"-->
<!-- #include virtual="/lib/db/db3close.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->

<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ����������������ǰ �������
' Hieditor : 2011.07.21 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/classes/offshop/offshopitemcls.asp"-->
<!-- #include virtual="/lib/classes/offshop/stock/shortagestock_cls.asp" -->
<!-- #include virtual="/lib/classes/items/adminshoppingbag/adminshoppingbag_cls.asp" -->
<%
dim page , shopid , isusing , makerid , itemid , itemname , generalbarcode , i , sell7days
dim cdl , cdm , cds , shortagetype , comm_cd ,includepreorder ,research , parameter , ipgo
    page = requestCheckVar(getNumeric(request("page")),10)
    research = requestCheckVar(request("research"),2)
    isusing = requestCheckVar(request("isusing"),1)
    makerid = requestCheckVar(request("makerid"),32)
    itemid = requestCheckVar(request("itemid"),10)
    itemname = requestCheckVar(request("itemname"),64)
    generalbarcode = requestCheckVar(request("generalbarcode"),20)
    comm_cd = requestCheckVar(request("comm_cd"),16)
    cdl = requestCheckVar(getNumeric(request("cdl")),3)
    cdm = requestCheckVar(getNumeric(request("cdm")),3)
    cds = requestCheckVar(getNumeric(request("cds")),3)
    shortagetype = requestCheckVar(request("shortagetype"),10)
    includepreorder = requestCheckVar(request("includepreorder"),2)
    sell7days = requestCheckVar(request("sell7days"),2)
    ipgo = requestCheckVar(request("ipgo"),2)
	shopid = requestCheckVar(request("shopid"),32)

if page="" then page=1
if (research<>"on") and (includepreorder="") then
    includepreorder = "on"
end if
if (research<>"on") and (ipgo="") then
    ipgo = "on"
end if
if (research<>"on") and (shortagetype="") then
    shortagetype = 7
end if
if (research<>"on") and (isusing="") then
    isusing = "Y"
end if

'/�����ϰ�� ���� ���常 ��밡��
if (C_IS_SHOP) then
	'/���α��� ���� �̸�
	if getlevel_sn("",session("ssBctId")) > 6 then
		shopid = C_STREETSHOPID
	end if
else
	if (C_IS_Maker_Upche) then
		shopid = session("ssBctID")
	else
		if (Not C_ADMIN_USER) then
			loginidshopormaker = "--"		'ǥ�þ��Ѵ�. ����.
		else

		end if
	end if
end if

if shopid = "" then shopid = "streetshop011"

parameter = "page="&page&"&research="&research&"&shopid="&shopid&"&isusing="&isusing&"&makerid="&makerid&"&itemid="&itemid&"&itemname="&itemname&"&sell7days="&sell7days&""
parameter = parameter & "&generalbarcode="&generalbarcode&"&comm_cd="&comm_cd&"&cdl="&cdl&"&cdm="&cdm&"&cds="&cds&"&shortagetype="&shortagetype&"&includepreorder="&includepreorder
parameter = parameter & "&ipgo="&ipgo&""

dim oshortage
set oshortage  = new cshortagestock_list
    oshortage.FPageSize = 100
    oshortage.FCurrPage = page
    oshortage.frectcdl = cdl
    oshortage.frectcdm = cdm
    oshortage.frectcds = cds
    oshortage.frectincludepreorder = includepreorder
    oshortage.frectsell7days = sell7days
    oshortage.Frectshopid = shopid
    oshortage.Frectisusing = isusing
    oshortage.Frectmakerid = makerid
    oshortage.Frectitemid = itemid
    oshortage.Frectitemname = itemname
    oshortage.Frectcomm_cd = comm_cd
    oshortage.Frectgeneralbarcode = generalbarcode
    oshortage.Frectshortagetype = shortagetype
    oshortage.Frectipgo = ipgo

    if shopid <> "" then
        oshortage.fshortagestock_list
    else
        response.write "<script language='javascript'>"
        response.write "    alert('������ ������ �ּ���');"
        response.write "</script>"
    end if

dim IsUpcheWitakItem
if (makerid<>"") and (shopid<>"") then
    IsUpcheWitakItem = (GetShopBrandContract(shopid,makerid)="B012")
end if

Response.Expires=0
response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment; filename=TEN" & Left(CStr(now()),10) & "_����������������ǰ.xls"
Response.CacheControl = "public"
%>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
    <td colspan="25">
        �˻���� : <b><%= oshortage.FTotalcount %></b>
        &nbsp;
        <b><%= page %> / <%= oshortage.FTotalpage %></b>
    </td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    <td>����</td>
    <td>
    	����ó
    </td>
    <td>�귣��</td>
    <td>��ǰ�ڵ�</td>
    <td>�̹���</td>
    <td>��ǰ��<br>[�ɼǸ�]</td>
	<% if not(C_IS_Maker_Upche) and not(C_IS_SHOP) then %>
    	<td>���԰�</td>
    <% end if %>
    <td>����<br>���ް�</td>
    <td>�ǸŰ�</td>
    <td>�����԰�<br>��ǰ</td>
    <td>�귣���԰�<br>��ǰ</td>
    <td>���Ǹ�<br>��Ȳ</td>
    <td>�ý���<br>���</td>
    <td>�ǻ�<br>����</td>
    <td>����</td>
    <td>��ȿ���</td>
    <td>�Ǹż���(3��)<br>�Ǹż���(7��)</td>
    <td>
        �ʿ����(3��)
        <br>�ʿ����(7��)
        <br>�ʿ����(14��)
        <!--<br>�ʿ����(28��)-->
    </td>
    <td>����</td>
    <td>���</td>
</tr>
<% if oshortage.FresultCount > 0 then %>
<% for i=0 to oshortage.FresultCount -1 %>
<% if oshortage.FItemlist(i).Fisusing="N" then %>
<tr bgcolor="#EEEEEE" align="center">
<% else %>
<tr bgcolor="#FFFFFF" align="center">
<% end if %>

    <td >
        <%= oshortage.FItemlist(i).fshopid %>
    </td>
    <td>
        <!--<%'= mwdivName(oshortage.FItemlist(i).fcentermwdiv) %><p>-->
        <%= GetdeliverGubunName(oshortage.FItemlist(i).fcomm_cd) %><br>(<%= GetJungsanGubunName(oshortage.FItemlist(i).fcomm_cd) %>)
    </td>
    <td>
        <a href="javascript:searchmakerid('<%= oshortage.FItemlist(i).fmakerid %>');" onfocus="this.blur()"><%= oshortage.FItemlist(i).fmakerid %></a>
    </td>
    <td>
        <%= oshortage.FItemlist(i).Fitemgubun %><%= CHKIIF(oshortage.FItemlist(i).Fitemid>=1000000,Format00(8,oshortage.FItemlist(i).Fitemid),Format00(6,oshortage.FItemlist(i).Fitemid)) %><%= oshortage.FItemlist(i).Fitemoption %>
    </td>
    <td>
        <img src="<%= oshortage.FItemlist(i).GetImageSmall %>" width=50 height=50 border=0>
    </td>
    <td>
        <%= oshortage.FItemlist(i).fshopitemname %><Br>
        <% if oshortage.FItemlist(i).fshopitemoptionname <> "" then %>
            [<%=oshortage.FItemlist(i).fshopitemoptionname%>]
        <% end if %>
    </td>
	<% if not(C_IS_Maker_Upche) and not(C_IS_SHOP) then %>
	    <td>
	        <%= FormatNumber(oshortage.FItemlist(i).fshopsuplycash,0) %>
	    </td>
	<% end if %>
    <td>
        <%= FormatNumber(oshortage.FItemlist(i).fshopbuyprice,0) %>
    </td>
    <td>
        <%= FormatNumber(oshortage.FItemlist(i).fshopitemprice,0) %>
    </td>
    <td>
        <%= oshortage.FItemlist(i).flogicsipgono + oshortage.FItemlist(i).flogicsreipgono %>    <!--�����԰��ǰ-->
    </td>
    <td>
        <%= oshortage.FItemlist(i).fbrandipgono + oshortage.FItemlist(i).fbrandreipgono %>		<!--�귣���԰��ǰ-->
    </td>
    <td>
        <%= oshortage.FItemlist(i).fsellno %>       <!--���Ǹ���Ȳ -->
    </td>
    <td>
        <%= oshortage.FItemlist(i).fsysstockno %>       <!--�ý������-->
    </td>
    <td>
        <%= oshortage.FItemlist(i).Ferrrealcheckno %>       <!--����-->
    </td>
    <td>
        <%= oshortage.FItemlist(i).ferrsampleitemno %>      <!--����-->
    </td>
    <td>
        <%= oshortage.FItemlist(i).getAvailStock %>     <!--��ȿ���-->
    </td>
    <td>
        <%= oshortage.FItemlist(i).fsell3days %> (3��)
        <p><%= oshortage.FItemlist(i).fsell7days %> (7��)      <!--�Ǹż���-->
    </td>
    <td>
        <!-- ���ʿ���� -->
        <% if oshortage.FItemlist(i).frequire3daystock > 0 then %>
            <a href="javascript:inputiteno('<%= oshortage.FItemlist(i).frequire3daystock %>',frmBuyPrc<%= i %>);" onfocus="this.blur()">
            <font color="red"><%= oshortage.FItemlist(i).frequire3daystock %> (3��)</font>
            </a>
        <% else %>
            0 (3��)
        <% end if %>
        <% if oshortage.FItemlist(i).frequire7daystock > 0 then %>
            <a href="javascript:inputiteno('<%= oshortage.FItemlist(i).frequire7daystock %>',frmBuyPrc<%= i %>);" onfocus="this.blur()"><p>
            <font color="red"><%= oshortage.FItemlist(i).frequire7daystock %> (7��)</font>
            </a>
        <% else %>
            <p>0 (7��)
        <% end if %>
        <% if oshortage.FItemlist(i).frequire14daystock > 0 then %>
            <a href="javascript:inputiteno('<%= oshortage.FItemlist(i).frequire14daystock %>',frmBuyPrc<%= i %>);" onfocus="this.blur()"><p>
            <font color="red"><%= oshortage.FItemlist(i).frequire14daystock %> (14��)</font>
            </a>
        <% else %>
            <p>0 (14��)
        <% end if %>
        <!--<%' if oshortage.FItemlist(i).frequire28daystock > 0 then %>
            <a href="javascript:inputiteno('<%= oshortage.FItemlist(i).frequire28daystock %>',frmBuyPrc<%= i %>);" onfocus="this.blur()">
            <font color="red"><%'= oshortage.FItemlist(i).frequire28daystock %> (28��)</font>
            </a>
        <%' else %>
            <p>0 (28��)
        <%' end if %>-->
    </td>
    <td>

    </td>
    <td>
        <% if oshortage.FItemList(i).Fpreorderno>0 then %>
        	���ֹ�:
            <% if oshortage.FItemList(i).Fpreorderno<>oshortage.FItemList(i).Fpreordernofix then response.write CStr(oshortage.FItemList(i).Fpreorderno) + " -> " %>
        	<%= oshortage.FItemList(i).Fpreordernofix %><br>
        <% end if %>

    </td>
</tr>

<% next %>

<% else %>

<tr bgcolor="#FFFFFF">
    <td colspan="25" align="center">[�˻������ �����ϴ�.]</td>
</tr>
<% end if %>

</table>

<%
	set oshortage = nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->

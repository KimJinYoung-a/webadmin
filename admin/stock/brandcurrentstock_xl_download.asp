<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  �귣�庰�����Ȳ �����ޱ�
' History : 2020.12.28 ������ ����
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stock/summary_itemstockcls.asp"-->
<%
dim makerid, onoffgubun, mwdiv, ImgUsing, sellyn, usingyn, danjongyn,isusing, limitrealstock,centermwdiv
dim returnitemgubun, itemname, itemidArr, cdl, cdm, cds, page, i, osummarystockbrand
dim stocktype
dim ordby
Dim dispCate : dispCate = RequestCheckvar(Request("disp"),10)
	makerid     = requestCheckvar(request("makerid"),32)
	onoffgubun  = requestCheckvar(request("onoffgubun"),9)
	ImgUsing    = requestCheckvar(request("ImgUsing"),9)
	sellyn      = requestCheckvar(request("sellyn"),9)
	usingyn     = requestCheckvar(request("usingyn"),9)
	danjongyn   = requestCheckvar(request("danjongyn"),9)
	mwdiv       = requestCheckvar(request("mwdiv"),9)
	returnitemgubun = requestCheckvar(request("returnitemgubun"),9)
	itemname        = requestCheckvar(request("itemname"),64)
	itemidArr       = Trim(requestCheckvar(request("itemidArr"),255))
	page            = requestCheckvar(request("page"),9)
	cdl = requestCheckvar(request("cdl"),3)
	cdm = requestCheckvar(request("cdm"),3)
	cds = requestCheckvar(request("cds"),3)
	limitrealstock 	= requestCheckvar(request("limitrealstock"),10)
	centermwdiv    	= requestCheckvar(request("centermwdiv"),10)
	ordby    		= requestCheckvar(request("ordby"),64)

	stocktype    	= requestCheckvar(request("stocktype"),32)
	if (stocktype = "") then
		stocktype = "sys"
	end if

if onoffgubun="" then onoffgubun="on"
if ImgUsing="" then ImgUsing="Y"
if Right(itemidArr,1)="," then itemidArr=Left(itemidArr,Len(itemidArr)-1)
if (page="") then page=1

set osummarystockbrand = new CSummaryItemStock
	osummarystockbrand.FPageSize = 5000
	osummarystockbrand.FCurrPage = page
	osummarystockbrand.FRectCD1   = cdl
	osummarystockbrand.FRectCD2   = cdm
	osummarystockbrand.FRectCD3   = cds
	osummarystockbrand.FRectItemIdArr = itemidArr
	osummarystockbrand.FRectItemName = itemname
	osummarystockbrand.FRectMakerid = makerid
	osummarystockbrand.FRectOnlySellyn = sellyn
	osummarystockbrand.FRectOnlyIsUsing = usingyn
	osummarystockbrand.FRectDanjongyn =danjongyn
	osummarystockbrand.FRectMWDiv = mwdiv
	osummarystockbrand.FRectCenterMWDiv = centermwdiv
	osummarystockbrand.FRectReturnItemGubun = returnitemgubun
	osummarystockbrand.FRectlimitrealstock = limitrealstock
	osummarystockbrand.FRectStockType = stocktype
	osummarystockbrand.FRectOrderBy = ordby
	osummarystockbrand.FRectDispCate = dispCate

	if (onoffgubun = "on") and ((itemidArr<>"") or (itemname<>"") or (makerid<>"") or (cdl<>"") or (mwdiv<>"")) then
		osummarystockbrand.GetCurrentStockByOnlineBrand
	elseif (Left(onoffgubun,3) = "off") then
		osummarystockbrand.FRectItemGubun =  Mid(onoffgubun,4,2)
		osummarystockbrand.GetCurrentStockByOfflineBrand
	end if

	response.ContentType = "application/vnd.ms-excel"
	Response.AddHeader "Content-Disposition", "attachment; filename=brandStock_itemlist_" & GetCurrentTimeFormat & ".xls"
%>
<html xmlns:o="urn:schemas-microsoft-com:office:office"
xmlns:x="urn:schemas-microsoft-com:office:excel"
xmlns="http://www.w3.org/TR/REC-html40">

<head>
<meta http-equiv=Content-Type content="text/html; charset=ks_c_5601-1987">
<meta name=ProgId content=Excel.Sheet>
<meta name=Generator content="Microsoft Excel 12">
<style type="text/css">
 td {font-size:8.0pt;}
 .txt {mso-number-format:"\@";}
 .num {mso-number-format:"0";}
 .prc {mso-number-format:"\#\,\#\#0";}
</style>
</head>

<body>
<!--[if !excel]>����<![endif]-->
<div align=center x:publishsource="Excel">
<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>��ǰ����</td>
	<td>��ǰ�ڵ�</td>
	<td>�ɼ��ڵ�</td>
	<td>��ǰ���ڵ�</td>
	<td>�������ڵ�</td>
	<% if ImgUsing<>"N" then %>
		<td width="50">�̹���</td>
	<% end if %>
	<td>�귣��ID</td>
	<td>��ǰ��</td>
	<td>�ɼǸ�</td>
	<td>��۱���</td>
	<td>���԰�/��ǰ</td>
	<td>ON�Ǹ�/��ǰ</td>
    <td>OFF���/��ǰ</td>
    <td>��Ÿ���/��ǰ</td>
    <td>CS���/��ǰ</td>
    <td bgcolor="F4F4F4">�ý��������</td>
    <td>�Ѻҷ�</td>
    <td>�ѿ���</td>
    <td>ON��ǰ�غ�</td>
    <td>OFF��ǰ�غ�</td>
    <td bgcolor="F4F4F4">����ľ����</td>
    <td>����ľ�</td>

	<td>�Ǹſ���</td>
	<td>��������</td>
	<td>��������</td>
	<td>��������</td>
	<% if ImgUsing="N" then %>
	<td>���ǸŰ�</td>
	<td>�����԰�</td>
	<% end if %>
</tr>
<% for i=0 to osummarystockbrand.FResultCount - 1 %>
    <% if osummarystockbrand.FItemList(i).Fisusing="Y" then %>
    <tr bgcolor="#FFFFFF" align="center" <%=chkIIF(ImgUsing<>"N","height=""52""","")%>>
    <% else %>
    <tr bgcolor="#EEEEEE" align="center" <%=chkIIF(ImgUsing<>"N","height=""52""","")%>>
    <% end if %>
    	<td class="txt"><%= osummarystockbrand.FItemList(i).Fitemgubun %></td>
    	<td class="num"><%= osummarystockbrand.FItemList(i).Fitemid %></td>
    	<td class="txt"><%= osummarystockbrand.FItemList(i).Fitemoption %></td>
    	<td class="txt"><%= osummarystockbrand.FItemList(i).Fitemrackcode %></td>
		<td class="txt"><%= osummarystockbrand.FItemList(i).Fsubitemrackcode %></td>
    	<% if ImgUsing<>"N" Then %>
    	<td><img src="<%= osummarystockbrand.FItemList(i).Fimgsmall %>" width=50 height=50> </td>
    	<% end if %>
    	<td class="txt"><%= osummarystockbrand.FItemList(i).FMakerid %></td>
    	<td class="txt"><%= osummarystockbrand.FItemList(i).Fitemname %></td>
        <td class="txt"><%= osummarystockbrand.FItemList(i).FitemoptionName %></td>
        <td class="txt"><%= fnColor(osummarystockbrand.FItemList(i).Fmwdiv,"mw") %></td>
    	<td class="prc"><%= osummarystockbrand.FItemList(i).Ftotipgono %></td>
    	<td class="prc"><%= -1*osummarystockbrand.FItemList(i).Ftotsellno %></td>
    	<td class="prc"><%= osummarystockbrand.FItemList(i).Foffchulgono + osummarystockbrand.FItemList(i).Foffrechulgono %></td>
    	<td class="prc"><%= osummarystockbrand.FItemList(i).Fetcchulgono + osummarystockbrand.FItemList(i).Fetcrechulgono %></td>
    	<td class="prc"><%= osummarystockbrand.FItemList(i).Ferrcsno %></td>
    	<td class="prc" bgcolor="F4F4F4"><b><%= osummarystockbrand.FItemList(i).Ftotsysstock %></b></td>
    	<td class="prc"><%= osummarystockbrand.FItemList(i).Ferrbaditemno %></td>
    	<td class="prc"><%= osummarystockbrand.FItemList(i).Ferrrealcheckno %></td>
    	<td class="prc"><%= osummarystockbrand.FItemList(i).Fipkumdiv5 %></td>
    	<td class="prc"><%= osummarystockbrand.FItemList(i).Foffconfirmno %></td>
    	<td class="prc" bgcolor="F4F4F4"><b><%= osummarystockbrand.FItemList(i).GetCheckStockNo %></b></td>
    	<td></td>
		<td class="txt"><%= fnColor(osummarystockbrand.FItemList(i).Fsellyn,"yn") %></td>
		<td class="txt"><%= fnColor(osummarystockbrand.FItemList(i).Flimityn,"yn") %></td>
		<td class="prc"><%= osummarystockbrand.FItemList(i).GetLimitStr %></td>
    	<td class="txt"><%= fnColor(osummarystockbrand.FItemList(i).Fdanjongyn,"dj") %></td>
        <% if ImgUsing="N" then %>
        <td class="prc"><%= osummarystockbrand.FItemList(i).FOnlineCurrentSellcash %></td>
        <td class="prc"><%= osummarystockbrand.FItemList(i).FOnlineCurrentBuycash %></td>
        <% end if %>
    </tr>
<% next %>
</table>
</div>
</body>
</html>
<%
set osummarystockbrand = Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->

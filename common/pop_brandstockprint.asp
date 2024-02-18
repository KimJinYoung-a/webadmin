<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  �귣�庰�����Ȳ
' History : 2009.04.07 ������ ����
'			2013.10.16 �ѿ�� ����
'			2019.11.07 ������ ���� (�������ڵ��߰�)
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stock/summary_itemstockcls.asp"-->
<%
dim makerid, onoffgubun, mwdiv, ImgUsing, sellyn, usingyn, danjongyn,isusing, limitrealstock,centermwdiv
dim returnitemgubun, itemname, itemidArr, cdl, cdm, cds, page, i, osummarystockbrand
dim stocktype, useoffinfo, itemgubun, startMon, endMon, excits, vPurchasetype
dim limityn, itemgrade, ordby, itemrackcode, bulkstockgubun, warehouseCd, agvstockgubun
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
	vPurchasetype 	= request("purchasetype")
    limityn  		= requestCheckvar(request("limityn"),2)
    itemgrade     	= RequestCheckVar(request("itemgrade"),32)
	stocktype    	= requestCheckvar(request("stocktype"),32)
	itemgubun     	= RequestCheckVar(request("itemgubun"),32)
	startMon     	= RequestCheckVar(request("startMon"),32)
	endMon     		= RequestCheckVar(request("endMon"),32)
	if (stocktype = "") then
		stocktype = "sys"
	end if
    itemrackcode    = RequestCheckVar(request("itemrackcode"),32)
    bulkstockgubun  = RequestCheckVar(request("bulkstockgubun"),32)
    warehouseCd  	= RequestCheckVar(request("warehouseCd"),32)
	useoffinfo = request("useoffinfo")
    excits  	= RequestCheckVar(request("excits"),32)
    agvstockgubun  	= RequestCheckVar(request("agvstockgubun"),32)

'if onoffgubun="" then onoffgubun="on"
if ImgUsing="" then ImgUsing="Y"
if Right(itemidArr,1)="," then itemidArr=Left(itemidArr,Len(itemidArr)-1)
if (page="") then page=1

if (onoffgubun = "") and (itemgubun = "") then
	itemgubun="10"
elseif (onoffgubun <> "") and (itemgubun = "") then
	if (onoffgubun = "on") then
		itemgubun="10"
	elseif (onoffgubun = "off") then
		itemgubun="exc10"
	else
		itemgubun = Right(onoffgubun,2)
	end if
end if
if itemgubun="" then itemgubun="10"

if itemgubun = "10" then
	onoffgubun = "on"
elseif (itemgubun = "exc10") then
	onoffgubun = "off"
elseif (itemgubun <> "10") then
	onoffgubun = "off" & itemgubun
end if

'//��ǰ�ڵ� ��ȿ�� �˻�
if itemidArr<>"" then
	dim iA ,arrTemp,arrItemid
  itemidArr = replace(itemidArr,chr(13),"")
	arrTemp = Split(itemidArr,chr(10))

	iA = 0
	do while iA <= ubound(arrTemp)
		if Trim(arrTemp(iA))<>"" and isNumeric(Trim(arrTemp(iA))) then
			arrItemid = arrItemid & Trim(arrTemp(iA)) & ","
		end if
		iA = iA + 1
	loop

	if len(arrItemid)>0 then
		itemidArr = left(arrItemid,len(arrItemid)-1)
	else
		if Not(isNumeric(itemidArr)) then
			itemidArr = ""
		end if
	end if
end if

set osummarystockbrand = new CSummaryItemStock
	osummarystockbrand.FPageSize = 1000
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
	osummarystockbrand.FRectDispCate = dispCate
    osummarystockbrand.FRectRackCode = itemrackcode
    osummarystockbrand.FRectBulkStockGubun = bulkstockgubun
    osummarystockbrand.FRectWarehouseCd = warehouseCd
	osummarystockbrand.FRectUseOffInfo = useoffinfo
    osummarystockbrand.FRectExcIts = excits
	osummarystockbrand.FRectPurchasetype = vPurchasetype
    osummarystockbrand.FRectLimitYN = limityn
    osummarystockbrand.FRectAgvStockGubun = agvstockgubun

	if (ordby = "1") then
		osummarystockbrand.FRectOrderBy = "T.itemid desc"
	elseif (ordby = "2") then
		osummarystockbrand.FRectOrderBy = "T.itemrackcode asc,T.itemid desc"
	end if

	if IsNumeric(startMon) then
		osummarystockbrand.FRectStartDate = startMon
	elseif (startMon <> "") then
		response.write "<script>alert('������ ���ڸ� �����մϴ�. " & startMon & "')</script>"
	end if
	if IsNumeric(endMon) then
		osummarystockbrand.FRectEndDate = endMon
	elseif (endMon <> "") then
		response.write "<script>alert('������ ���ڸ� �����մϴ�. " & endMon & "')</script>"
	end if

	if (onoffgubun = "on") and ((itemidArr<>"") or (itemname<>"") or (makerid<>"") or (cdl<>"") or (mwdiv<>"")) then
		osummarystockbrand.GetCurrentStockByOnlineBrandNEW
	elseif (Left(onoffgubun,3) = "off") then
		osummarystockbrand.FRectItemGubun =  Mid(onoffgubun,4,2)
		osummarystockbrand.GetCurrentStockByOfflineBrand
	end if
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript">

function NextPage(page){
    document.frm.page.value = page;
    document.frm.submit();
}

function PopItemSellEdit(iitemid){
	var popwin = window.open('/admin/lib/popitemsellinfo.asp?itemid=' + iitemid,'itemselledit','width=500,height=600,scrollbars=yes,resizable=yes')
}

function RefreshPageByImg(mode){
	document.noImgFrm.ImgUsing.value=mode;
	document.noImgFrm.submit();
}

function RefreshPageByOrdBy() {
	document.noImgFrm.ImgUsing.value='<%= ImgUsing %>';
	document.noImgFrm.ordby.value=document.frm.ordby.value;
	document.noImgFrm.submit();
}

function chkEnDisabled(comp){
    var frm = comp.form;
    if (comp.value==""){
       frm.sellyn.disabled=false;
       //frm.usingyn.disabled=false;
       frm.danjongyn.disabled=false;
    }else{
       frm.sellyn.disabled=true;
       //frm.usingyn.disabled=true;
       frm.danjongyn.disabled=true;
    }
}

function DownloadBrandStockXLS(){
	var onoffgubun = document.noImgFrm.onoffgubun.value;
    var makerid = document.noImgFrm.makerid.value;
	var mwdiv = document.noImgFrm.mwdiv.value;
	var centermwdiv = document.noImgFrm.centermwdiv.value;
    var sellyn = document.noImgFrm.sellyn.value;
    var isusing= document.noImgFrm.isusing.value;
	var danjongyn  = document.noImgFrm.danjongyn.value;
	var disp     = document.noImgFrm.disp.value;
    var itemidArr     = document.noImgFrm.itemidArr.value.replace(/(?:\r\n|\r|\n)/g, ',');
    var itemname     = document.noImgFrm.itemname.value.replace(/(?:\r\n|\r|\n)/g, ',');
    var limitrealstock     = document.noImgFrm.limitrealstock.value;
	var stocktype     = document.noImgFrm.stocktype.value;
	var returnitemgubun = document.noImgFrm.returnitemgubun.value;
	var ImgUsing     = document.noImgFrm.ImgUsing.value;
	var ordby     = document.noImgFrm.ordby.value;

	self.location.href='/admin/stock/brandcurrentstock_xl_download.asp?makerid=' + makerid + '&stocktype=' + stocktype + '&itemidArr=' + itemidArr + '&disp=' + disp + '&onoffgubun=' + onoffgubun + '&mwdiv=' + mwdiv + '&centermwdiv=' + centermwdiv + '&sellyn=' + sellyn + '&isusing=' + isusing + '&danjongyn=' + danjongyn + '&returnitemgubun=' + returnitemgubun + '&itemname=' + itemname + '&limitrealstock=' + limitrealstock +"&ImgUsing="+ImgUsing+"&ordby="+ordby,'pop_brandstockprint','width=1000,height=600,scrollbars=yes,resizable=yes';
}
</script>

<!-- ǥ ��ܹ� ����-->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="research" value="on">
<input type="hidden" name="page" value="">
<tr height="25" bgcolor="<%= adminColor("tabletop") %>">
	<td colspan="3">
		<img src="/images/icon_star.gif" align="absbottom">
		<font color="red"><strong>����ľ� SHEET ��� <font color="#000000">[<%= makerid %>]</font></strong></font>
	</td>
</tr>
<tr height="25" bgcolor="#FFFFFF">
        <td>
			* �귣��: <% drawSelectBoxDesignerwithName "makerid", makerid %>
			&nbsp;&nbsp;
			* ��ǰ��: <input type="text" class="text" name="itemname" value="<%= itemname %>" size="32" maxlength="32">
			&nbsp;&nbsp;
			<!-- #include virtual="/common/module/dispCateSelectBox.asp"-->
			&nbsp;&nbsp;
			* ��ǰ�ڵ�: <textarea rows="3" cols="10" name="itemidArr" id="itemidArr"><%=replace(itemidArr,",",chr(10))%></textarea>
			<input type=checkbox name="useoffinfo" <% if useoffinfo = "on" then response.write "checked" %> > ������ǰ(10) ����(OFF��ǰ �˻���)
			<br>
			<!--
			<select name="onoffgubun" >
				<option value="on" <%= ChkIIF(onoffgubun="on","selected","") %> >ON��ǰ</option>
				<option value="off" <%= ChkIIF(onoffgubun="off","selected","") %> >OFF��ǰ</option>
				<option value="off70" <%= ChkIIF(onoffgubun="off70","selected","") %> >OFF��ǰ-70</option>
				<option value="off80" <%= ChkIIF(onoffgubun="off80","selected","") %> >OFF��ǰ-80</option>
                <option value="off85" <%= ChkIIF(onoffgubun="off85","selected","") %> >OFF��ǰ-85</option>
				<option value="off90" <%= ChkIIF(onoffgubun="off90","selected","") %> >OFF��ǰ-90</option>
			</select>
			-->
			<input type="hidden" name="onoffgubun" value="<%= onoffgubun %>">
			* ��ǰ����: <% drawSelectBoxItemGubunForSearch "itemgubun", itemgubun %>
			&nbsp;&nbsp;
			* �Ǹ�:<% drawSelectBoxSellYN "sellyn", sellyn %>
			&nbsp;&nbsp;
			* ���:<% drawSelectBoxUsingYN "usingyn", usingyn %>
			&nbsp;&nbsp;
			* ����:<% drawSelectBoxDanjongYN "danjongyn", danjongyn %>
			<br>
			* �ŷ����� :<% drawSelectBoxMWU "mwdiv", mwdiv %>
			&nbsp;&nbsp;
			* ���͸��Ա��� :
    		<select class="select" name="centermwdiv">
            <option value="">��ü</option>
            <option value="M" <% if centermwdiv="M" then response.write "selected" %> >����</option>
            <option value="W" <% if centermwdiv="W" then response.write "selected" %> >��Ź</option>
            <option value="N" <% if centermwdiv="N" then response.write "selected" %> >������</option>
            </select>
            &nbsp;&nbsp;
			* �������� : 
			<% drawPartnerCommCodeBox true,"purchasetype","purchasetype",vPurchasetype,"" %>
			&nbsp;&nbsp;
			<span style="white-space:nowrap;">����:<% drawSelectBoxLimitYN "limityn", limityn %></span>
			<br>
			* ��ǰ��� :
			<select class="select" name="itemgrade">
				<option value="">��ü</option>
				<option value="A" <% if itemgrade="A" then response.write "selected" %> >A</option>
				<option value="B" <% if itemgrade="B" then response.write "selected" %> >B</option>
				<option value="C" <% if itemgrade="C" then response.write "selected" %> >C</option>
				<option value="Z" <% if itemgrade="Z" then response.write "selected" %> >Z</option>
				<option value="AB" <% if itemgrade="AB" then response.write "selected" %> >A+B</option>
				<option value="ABC" <% if itemgrade="ABC" then response.write "selected" %> >A+B+C</option>
			</select>
			&nbsp;&nbsp;
			* ��ũ��� :
			<select class="select" name="bulkstockgubun">
				<option value="">��ü</option>
				<option value="nul" <% if bulkstockgubun="nul" then response.write "selected" %> >�Է�����</option>
				<option value="err" <% if bulkstockgubun="err" then response.write "selected" %> >��ũ���� ����</option>
			</select>
            &nbsp;&nbsp;
		    * �������� :
            <select class="select" name="warehouseCd">
                <option value="">��ü</option>
                <option value="AGV" <% if warehouseCd="AGV" then response.write "selected" %> >AGV</option>
                <option value="BLK" <% if warehouseCd="BLK" then response.write "selected" %> >��ũ</option>
            </select>
			&nbsp;&nbsp;
			* AGV��� :
			<select class="select" name="agvstockgubun">
				<option value="">��ü</option>
				<option value="availdiff" <% if agvstockgubun="availdiff" then response.write "selected" %> >��ȿ��� ����ġ��</option>
				<option value="ipkum5diff" <% if agvstockgubun="ipkum5diff" then response.write "selected" %> >����ľ���� ����ġ��</option>
				<option value="oneup" <% if agvstockgubun="oneup" then response.write "selected" %> >1�̻�</option>
				<option value="zero" <% if agvstockgubun="zero" then response.write "selected" %> >0</option>
				<option value="minus" <% if agvstockgubun="minus" then response.write "selected" %> >���̳ʽ�</option>
			</select>
        </td>
        <td align="right">
            <input type="button" class="button" name="refresh" value="���ΰ�ħ" onclick="document.frm.submit();">
        	<% if ImgUsing="Y" then %>
        		<input type="button" class="button" name="brandstockprint" value="�̹������ֱ�" onclick="RefreshPageByImg('N');">
        	<% else %>
        		<input type="button" class="button" name="brandstockprint" value="�̹������̱�" onclick="RefreshPageByImg('Y');">
        	<% end if %>
        	<input type="button" class="button" name="brandstockprint" value="����ϱ�" onclick="window.print();">
        </td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left" colspan="2">
	    * ���������� :
	    <input type=radio name="returnitemgubun" value="" <% if returnitemgubun="" then response.write "checked" %> onClick="chkEnDisabled(this);">��ü
		<input type=radio name="returnitemgubun" value="rackdisp" <% if returnitemgubun="rackdisp" then response.write "checked" %> onClick="chkEnDisabled(this);">������ ��ǰ [(�Ǹ�<>'N') or (�����ƴ�)]
		<input type=radio name="returnitemgubun" value="reton" <% if returnitemgubun="reton" then response.write "checked" %> onClick="chkEnDisabled(this);">��ǰ��� ��ǰ [(�Ǹ�='N') and (����) and (�ǻ���ȿ���<>0)]
	    <input type=radio name="returnitemgubun" value="retfin" <% if returnitemgubun="retfin" then response.write "checked" %> onClick="chkEnDisabled(this);">��ǰ�Ϸ� ��ǰ [(�Ǹ�='N') and (����) and (�ǻ���ȿ���=0)]
	    <script language='javascript'>chkEnDisabled(frm.returnitemgubun[<%= ChkIIF(returnitemgubun="","0",ChkIIF(returnitemgubun="rackdisp","1","2")) %>]);</script>
	    <p>
	    * <select name="stocktype" class="select">
			<option value="sys" <% if (stocktype = "sys") then %>selected<% end if %> >�ý������</option>
			<option value="real" <% if (stocktype = "real") then %>selected<% end if %> >��ȿ���</option>
		</select>
		: <% drawSelectBoxexistsstock "limitrealstock", limitrealstock, "" %>
		&nbsp;&nbsp;
		* ������ :
		<input type="text" class="text" name="startMon" size="2" value="<%= startMon %>">
		~
		<input type="text" class="text" name="endMon" size="2" value="<%= endMon %>"> ����
		&nbsp;&nbsp;
		* ���ļ��� :
		<select class="select" name="ordby" onchange="RefreshPageByOrdBy();">
			<option value="1" <%= CHKIIF(ordby = "1", "selected", "") %> >��ǰ�ڵ�</option>
			<option value="2" <%= CHKIIF(ordby = "2", "selected", "") %> >���ڵ�</option>
		</select>
        &nbsp;&nbsp;
        <input type="checkbox" class="checkbox" name="excits" value="Y" <%= CHKIIF(excits="Y", "checked", "") %> > 3PL ����
		&nbsp;&nbsp;
		* ���ڵ� :
		<input type="text" class="text" name="itemrackcode" size="8" value="<%= itemrackcode %>">
	</td>
</tr>
</form>
<form name="noImgFrm" method="post" action="">
	<input type="hidden" name="makerid" value="<%= makerid %>">
	<input type="hidden" name="cdl" value="<%= cdl %>">
	<input type="hidden" name="cdm" value="<%= cdm %>">
	<input type="hidden" name="cds" value="<%= cds %>">
	<input type="hidden" name="disp" value="<%= dispCate %>">
	<input type="hidden" name="centermwdiv" value="<%= centermwdiv %>">
	<input type="hidden" name="returnitemgubun" value="<%= returnitemgubun %>">
	<input type="hidden" name="sellyn" value="<%= sellyn %>">
	<input type="hidden" name="itemidArr" value="<%= itemidArr %>">
	<input type="hidden" name="onoffgubun" value="<%= onoffgubun %>">
	<input type="hidden" name="mwdiv" value="<%= mwdiv %>">
	<input type="hidden" name="isusing" value="<%= isusing %>">
	<input type="hidden" name="danjongyn" value="<%= danjongyn %>">
	<input type="hidden" name="itemname" value="<%= itemname %>">
	<input type="hidden" name="limitrealstock" value="<%= limitrealstock %>">
	<input type="hidden" name="stocktype" value="<%= stocktype %>">
	<input type="hidden" name="ImgUsing" value="<%=ImgUsing%>">
	<input type="hidden" name="ordby" value="<%=ordby%>">
	<input type="hidden" name="page" value="<%= page %>">
</form>
</table>
<!-- ǥ ��ܹ� ��-->
<p style="text-align:right;"><img src="/images/btn_excel.gif" title="����ľ� EXCEL�ޱ�" style="cursor:pointer" onclick="DownloadBrandStockXLS();" /></p>
<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="30">
		�˻���� : <b><%= osummarystockbrand.FTotalCount %></b>
		&nbsp;
		������ :
		<b><%= page %> / <%= osummarystockbrand.FTotalPage %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="30">��ǰ<br>����</td>
	<td width="60">��ǰ�ڵ�</td>
	<td width="40">�ɼ�<br>�ڵ�</td>
	<td width="60">��ǰ���ڵ�<br>(����)</td>
	<% if ImgUsing<>"N" then %>
		<td width="50">�̹���</td>
	<% end if %>

	<td>�귣��ID</td>
	<td>��ǰ��</td>
	<td>�ɼǸ�</td>
	<td width="30">���<br>����</td>
	<td width="35">��<br>�԰�<br>��ǰ</td>
	<td width="35">ON<br>�Ǹ�<br>��ǰ</td>
    <td width="35">OFF<br>���<br>��ǰ</td>
    <td width="30">��Ÿ<br>���<br>��ǰ</td>
    <td width="30">CS<br>���<br>��ǰ</td>
    <td width="40" bgcolor="F4F4F4">�ý���<br>�����</td>
    <td width="30">��<br>�ҷ�</td>
    <td width="30">��<br>����</td>
    <td width="30">ON<br>��ǰ<br>�غ�</td>
    <td width="30">OFF<br>��ǰ<br>�غ�</td>

    <td width="30" bgcolor="F4F4F4">���<br>�ľ�<br>���</td>
    <td width="50">���<br>�ľ�</td>
    <td width="30">AGV<br />���</td>

	<td width="30">�Ǹ�<br>����</td>
	<td width="30">����<br>����</td>
	<td width="50">����<br>����</td>
	<% if ImgUsing="N" then %>
	<td width="30">���ǸŰ�</td>
	<td width="30">�����԰�</td>
	<% end if %>
<!--
    <td width="30">ON<br>����<br>�Ϸ�</td>
    <td width="30">ON<br>�ֹ�<br>����</td>
    <td width="30">OFF<br>�ֹ�<br>����</td>
    <td width="30" bgcolor="F4F4F4">����<br>���</td>
    <td width="40" bgcolor="F4F4F4">�غ���<br>���</td>
-->
</tr>
<% for i=0 to osummarystockbrand.FResultCount - 1 %>
    <% if osummarystockbrand.FItemList(i).Fisusing="Y" then %>
    <tr bgcolor="#FFFFFF" align="center">
    <% else %>
    <tr bgcolor="#EEEEEE" align="center">
    <% end if %>
    	<td>
    		<%= osummarystockbrand.FItemList(i).Fitemgubun %>
    	</td>
    	<td><a href="javascript:PopItemSellEdit('<%= osummarystockbrand.FItemList(i).Fitemid %>');"><%= osummarystockbrand.FItemList(i).Fitemid %></a></td>
    	<td>
    		<%= osummarystockbrand.FItemList(i).Fitemoption %>
    	</td>
    	<td>
			<%= osummarystockbrand.FItemList(i).Fitemrackcode %>
			<% if osummarystockbrand.FItemList(i).Fsubitemrackcode <> "" then %>
				<br>(<%= osummarystockbrand.FItemList(i).Fsubitemrackcode %>)
			<% end if %>
		</td>
    	<% if ImgUsing<>"N" Then %>
    	<td><img src="<%= osummarystockbrand.FItemList(i).Fimgsmall %>" width=50 height=50> </td>
    	<% end if %>
    	<td><%= osummarystockbrand.FItemList(i).FMakerid %></td>
    	<td align="left">
        	<%= osummarystockbrand.FItemList(i).Fitemname %><br />
            [<%= osummarystockbrand.FItemList(i).FpublicBarcode %>]
        </td>
        <td align="left">
          	<%= osummarystockbrand.FItemList(i).FitemoptionName %>
        </td>
        <td><%= fnColor(osummarystockbrand.FItemList(i).Fmwdiv,"mw") %></td>
    	<td><%= osummarystockbrand.FItemList(i).Ftotipgono %></td>
    	<td><%= -1*osummarystockbrand.FItemList(i).Ftotsellno %></td>
    	<td><%= osummarystockbrand.FItemList(i).Foffchulgono + osummarystockbrand.FItemList(i).Foffrechulgono %></td>
    	<td><%= osummarystockbrand.FItemList(i).Fetcchulgono + osummarystockbrand.FItemList(i).Fetcrechulgono %></td>
    	<td><%= osummarystockbrand.FItemList(i).Ferrcsno %></td>
    	<td bgcolor="F4F4F4"><b><%= osummarystockbrand.FItemList(i).Ftotsysstock %></b></td>
    	<td><%= osummarystockbrand.FItemList(i).Ferrbaditemno %></td>
    	<td><%= osummarystockbrand.FItemList(i).Ferrrealcheckno %></td>
    	<td><%= osummarystockbrand.FItemList(i).Fipkumdiv5 %></td>
    	<td><%= osummarystockbrand.FItemList(i).Foffconfirmno %></td>

    	<td bgcolor="F4F4F4"><b><%= osummarystockbrand.FItemList(i).GetCheckStockNo %></b></td>
    	<td></td>
        <% if osummarystockbrand.FItemList(i).Fagvstock = 0 then %>
        <td></td>
        <% else %>
        <td bgcolor="F4F4F4"><b><%= osummarystockbrand.FItemList(i).Fagvstock %></b></td>
        <% end if %>

		<td><%= fnColor(osummarystockbrand.FItemList(i).Fsellyn,"yn") %></td>
		<td>
    		<%= fnColor(osummarystockbrand.FItemList(i).Flimityn,"yn") %>
    		<% if osummarystockbrand.FItemList(i).Flimityn="Y" then %>
    		<br>(<%= osummarystockbrand.FItemList(i).GetLimitStr %>)
    		<% end if %>
    	</td>
    	<td><%= fnColor(osummarystockbrand.FItemList(i).Fdanjongyn,"dj") %></td>
        <% if ImgUsing="N" then %>
        <td><%= FormatNumber(osummarystockbrand.FItemList(i).FOnlineCurrentSellcash,0) %></td>
        <td><%= FormatNumber(osummarystockbrand.FItemList(i).FOnlineCurrentBuycash,0) %></td>
        <% end if %>
<!--
    	<td><%= osummarystockbrand.FItemList(i).Fipkumdiv4 %></td>
    	<td><%= osummarystockbrand.FItemList(i).Fipkumdiv2 %></td>
    	<td><%= osummarystockbrand.FItemList(i).Foffjupno %></td>
    	<td bgcolor="F4F4F4"><b><%= osummarystockbrand.FItemList(i).Frealstock + osummarystockbrand.FItemList(i).Fipkumdiv5 + osummarystockbrand.FItemList(i).Foffconfirmno + osummarystockbrand.FItemList(i).Fipkumdiv4 + osummarystockbrand.FItemList(i).Fipkumdiv2 + osummarystockbrand.FItemList(i).Foffjupno%></b></td>
    	<td>-</td>
-->
    </tr>
<% next %>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="25" align="center">
		<% if osummarystockbrand.HasPreScroll then %>
		<a href="javascript:NextPage('<%= osummarystockbrand.StartScrollPage-1 %>')">[pre]</a>
		<% else %>
			[pre]
		<% end if %>

		<% for i=0 + osummarystockbrand.StartScrollPage to osummarystockbrand.FScrollCount + osummarystockbrand.StartScrollPage - 1 %>
			<% if i>osummarystockbrand.FTotalpage then Exit for %>
			<% if CStr(page)=CStr(i) then %>
			<font color="red">[<%= i %>]</font>
			<% else %>
			<a href="javascript:NextPage('<%= i %>')">[<%= i %>]</a>
			<% end if %>
		<% next %>

		<% if osummarystockbrand.HasNextScroll then %>
			<a href="javascript:NextPage('<%= i %>')">[next]</a>
		<% else %>
			[next]
		<% end if %>
	</td>
</tr>
</table>

<script>
	document.frm.onoffgubun[0].disabled = true;
	document.frm.onoffgubun[1].disabled = true;
	document.frm.mwdiv[0].disabled = true;
	document.frm.mwdiv[1].disabled = true;
	document.frm.mwdiv[2].disabled = true;
</script>

<%
set osummarystockbrand = Nothing
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->

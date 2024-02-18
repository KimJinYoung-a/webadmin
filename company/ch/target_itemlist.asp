<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/company/incSessionCompany.asp" -->
<!-- #include virtual="/company/ch/incGlobalVariable.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/items/overseas/overseasCls.asp"-->
<%

dim itemid, itemname, makerid, sellyn, usingyn, mwdiv, limityn, overSeaYn, weightYn, itemrackcode, vRegUserID, vIsReg
dim cdl, cdm, cds, sortDiv, sortDiv2, sellcash1, sellcash2
dim page

itemid		= request("itemid")
itemname	= request("itemname")
makerid		= request("makerid")
sellyn		= request("sellyn")
usingyn		= request("usingyn")
mwdiv		= request("mwdiv")
limityn		= request("limityn")
overSeaYn	= request("overSeaYn")
weightYn	= request("weightYn")
itemrackcode= request("itemrackcode")
sortDiv		= request("sortDiv")
sortDiv2	= request("sortDiv2")
vRegUserID	= request("reguserid")
vIsReg		= request("isreg")
sellcash1	= request("sellcash1")
sellcash2	= request("sellcash2")

cdl = request("cdl")
cdm = request("cdm")
cds = request("cds")

page = request("page")

'�⺻��
if (page="") then page=1
if mwdiv="" then mwdiv="MW"
if overSeaYn="" then overSeaYn="Y"
if weightYn="" then weightYn="Y"
if sortDiv="" then sortDiv="new"
if sortDiv2="" then sortDiv2="weightup"


if itemid<>"" then
	dim iA ,arrTemp,arrItemid

	arrTemp = Split(itemid,",")

	iA = 0
	do while iA <= ubound(arrTemp)
		if Trim(arrTemp(iA))<>"" and isNumeric(Trim(arrTemp(iA))) then
			arrItemid = arrItemid & Trim(arrTemp(iA)) & ","
		end if
		iA = iA + 1
	loop

	if len(arrItemid)>0 then
		itemid = left(arrItemid,len(arrItemid)-1)
	else
		if Not(isNumeric(itemid)) then
			itemid = ""
		end if
	end if
end if


'==============================================================================
dim oitem

set oitem = new COverSeasItem

oitem.FPageSize         = 30
oitem.FCurrPage         = page
oitem.FRectMakerid      = makerid
oitem.FRectItemid       = itemid
oitem.FRectItemName     = itemname

oitem.FRectSellYN       = sellyn
oitem.FRectIsUsing      = usingyn
oitem.FRectLimityn      = limityn
oitem.FRectMWDiv        = mwdiv
oitem.FRectIsOversea	= overSeaYn
oitem.FRectIsWeight		= weightYn
oitem.FRectRackcode		= itemrackcode

oitem.FRectCate_Large   = cdl
oitem.FRectCate_Mid     = cdm
oitem.FRectCate_Small   = cds
oitem.FRectSortDiv		= sortDiv
oitem.FRectSortDiv2		= sortDiv2

oitem.FRectRegUserID	= vRegUserID
oitem.FRectIsReg		= vIsReg
oitem.FRectSellcash1	= sellcash1
oitem.FRectSellcash2	= sellcash2

oitem.GetOverSeasTargetItemList

dim i

%>

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<script language="JavaScript" src="/js/xl.js"></script>
<script language="JavaScript" src="/js/common.js"></script>
<script language="JavaScript" src="/js/report.js"></script>
<link rel="stylesheet" href="/css/scm.css" type="text/css">
<script>
function NextPage(ipage){
	document.frm.page.value= ipage;
	document.frm.submit();
}

function chgSort(srt){
	document.frm.sortDiv.value= srt;
	document.frm.submit();
}

function chgReg(reg){
	document.frm.isreg.value= reg;
	document.frm.submit();
}

// ============================================================================
// �ɼǼ��� -��ü
function editItemOption(itemid) {
	var param = "itemid=" + itemid;

	popwin = window.open('/common/pop_itemoption.asp?' + param ,'editItemOption','width=900,height=600,scrollbars=yes,resizable=yes');
	popwin.focus();
}

//�Ǹż���
function PopItemSellEdit(iitemid){
	var popwin = window.open('/common/pop_simpleitemedit.asp?itemid=' + iitemid,'itemselledit','width=500,height=600,scrollbars=yes,resizable=yes')
	popwin.focus();
}

// ============================================================================
// �̹�������
function editItemImage(itemid, makerid) {
	var param = "itemid=" + itemid;
	
	//if(makerid =="ithinkso"){
		//popwin = window.open('/common/pop_itemimage_ithinkso.asp?' + param ,'editItemImage','width=900,height=600,scrollbars=yes,resizable=yes');
	//}else{
		popwin = window.open('/common/pop_itemimage.asp?' + param ,'editItemImage','width=900,height=600,scrollbars=yes,resizable=yes');
	//}
	popwin.focus();
}

// ��ǰ���� �̹��� ���/����
function popItemContImage(itemid)
{
	var popwin = window.open("/admin/shopmaster/item_imgcontents_write.asp?mode=edit&itemid=" + itemid + "&menupos=423","popitemContImage","width=600 height=600 scrollbars=yes resizable=yes");
	popwin.focus();
}

// �����Ȳ �˾�
function PopItemStock(itemid){
	var popwin = window.open("/admin/stock/itemcurrentstock.asp?menupos=709&itemid=" + itemid,"popitemstocklist","width=1000 height=600 scrollbars=yes resizable=yes");
	popwin.focus();
}

// �⺻���� ����
function editItemBasicInfo(itemid) {
	var param = "itemid=" + itemid + "&makerid=<%= makerid %>&page=<%= page %>&menupos=<%= menupos %>";
	popwin = window.open('/admin/itemmaster/pop_ItemBasicInfo.asp?' + param ,'editItemBasic','width=750,height=600,scrollbars=yes,resizable=yes');
	popwin.focus();
}

// �ǸŰ� �� ���ް� ����
function editItemPriceInfo(itemid) {
	var param = "itemid=" + itemid + "&makerid=<%= makerid %>&page=<%= page %>&menupos=<%= menupos %>";
	popwin = window.open('/admin/itemmaster/pop_ItemPriceInfo.asp?' + param ,'editItemPrice','width=750,height=600,scrollbars=yes,resizable=yes');
	popwin.focus();
}


function PopItemWeightEdit(iitemid){
	var popwin = window.open('/warehouse/pop_ItemWeightEdit.asp?itembarcode=' + iitemid,'itemWeightEdit','width=500,height=300,scrollbars=yes,resizable=yes')
}

function PopItemContent(iitemid){
	var popwin = window.open('/admin/itemmaster/overseas/popItemContent.asp?countrycd=kr&itemid=' + iitemid,'itemWeightEdit','width=700,height=700,scrollbars=yes,resizable=yes')
}

function jsSearchBrandID(frmName,compName){
    var compVal = "";
    try{
        compVal = eval("document.all." + frmName + "." + compName).value;
    }catch(e){
        compVal = "";
    }

    var popwin = window.open("/company/ch/popBrandSearch.asp?frmName=" + frmName + "&compName=" + compName + "&rect=" + compVal,"popBrandSearch","width=800 height=400 scrollbars=yes resizable=yes");

	popwin.focus();
}

function itemlistXls()
{
	document.frm.action = "target_itemlist_xls.asp";
	document.frm.submit();
	
	document.frm.action = "target_itemlist.asp";
}
</script>
</head>
<body>
<table width="700" border="0" class="a">
	<tr>
		<td>&gt;&gt;�ǸŴ���ǰ����Ʈ</td>
	</tr>
</table>
<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method=get>
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="page" >
	<input type="hidden" name="sortDiv" value="<%=sortDiv%>">
	<input type="hidden" name="isreg" value="<%=vIsReg%>">
	<tr align="center" bgcolor="<%= adminColor("topbar") %>" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
		<td align="left">
			�귣�� :<%	drawSelectBoxDesignerWithName "makerid", makerid %>
			&nbsp;
			<!-- #include virtual="/common/module/categoryselectbox.asp"-->
			<br>
			��ǰ�ڵ� :
			<input type="text" class="text" name="itemid" value="<%= itemid %>" size="30" maxlength="100" onKeyPress="if (event.keyCode == 13) document.frm.submit();">(��ǥ�� �����Է°���)
			&nbsp;
			��ǰ�� :
			<input type="text" class="text" name="itemname" value="<%= itemname %>" size="32" maxlength="32">
		</td>
		
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="�˻�" onClick="javascript:document.frm.submit();">
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("topbar") %>" >
		<td align="left">
			�Ǹ�:<% drawSelectBoxSellYN "sellyn", sellyn %>
	     	&nbsp;
	     	���:<% drawSelectBoxUsingYN "usingyn", usingyn %>
	     	&nbsp;
		</td>
	</tr>
    </form>
</table>

<p>
<%
	If Request.ServerVariables("REMOTE_ADDR") = "61.252.133.15" Then
%>
<a href="javascript:itemlistXls();"><img src="http://webadmin.10x10.co.kr/images/btn_excel.gif" border="0"></a>
<br>
<%
	End If
%>

<!-- ����Ʈ ���� -->

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15">
			<table width="100%" cellpadding="0" cellspacing="0" class="a">
			<tr>
				<td>
					�˻���� : <b><%= oitem.FTotalCount%></b>
					&nbsp;
					������ : <b><%= page %> /<%=  oitem.FTotalpage %></b>
				</td>
				<td align="right">
					��Ͽ��� :
					<select name="reg" class="select" onchange="chgReg(this.value)">
						<option value="" <%= CHKIIF(vIsReg="","selected","") %>>��ü����</option>
						<option value="x" <%= CHKIIF(vIsReg="x","selected","") %>>�̵�ϸ�</option>
						<option value="o" <%= CHKIIF(vIsReg="o","selected","") %>>��ϸ�</option>
					</select>
					&nbsp;&nbsp;&nbsp;
					���Ĺ�� :
					<select name="sort" class="select" onchange="chgSort(this.value)">
						<option value="new" <% if sortDiv="new" then Response.Write "selected" %>>�Ż�ǰ��</option>
						<option value="best" <% if sortDiv="best" then Response.Write "selected" %>>�α��ǰ��</option>
						<option value="min" <% if sortDiv="min" then Response.Write "selected" %>>�������ݼ�</option>
						<option value="hi" <% if sortDiv="hi" then Response.Write "selected" %>>�������ݼ�</option>
						<option value="hs" <% if sortDiv="hs" then Response.Write "selected" %>>������������</option>
						<!--<option value="weight" <% if sortDiv="weight" then Response.Write "selected" %>>��ǰ���Լ�</option>//-->
					</select>
				</td>
			</tr>
			</table>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="60">No.</td>
		<td width=50> �̹���</td>
		<td width="100">�귣��ID</td>
		<td> ��ǰ��</td>
		<td width="60">�ǸŰ�</td>
		<td width="30">���<br>����</td>
		<td width="30">�Ǹ�<br>����</td>
		<td width="30">���<br>����</td>
		<td width="40">�ؿ�<br>����</td>
		<td width="60">��ǰ<br>����</td>
		<td width="100">��Ͽ���</td>
    </tr>
<% if oitem.FresultCount<1 then %>
    <tr bgcolor="#FFFFFF">
    	<td colspan="15" align="center">[�˻������ �����ϴ�.]</td>
    </tr>
<% end if %>
<% if oitem.FresultCount > 0 then %>
    <% for i=0 to oitem.FresultCount-1 %>
	<tr class="a" height="25" bgcolor="#FFFFFF">
		<td align="center">
			<a href="http://www.10x10.co.kr/shopping/category_prd.asp?itemid=<%= oitem.FItemList(i).Fitemid %>" target="_blank" title="�̸�����">				
			<%= oitem.FItemList(i).Fitemid %></a>
			</td>
		<td align="center"><img src="<%= oitem.FItemList(i).FSmallImage %>" width="50" height="50" border="0"></td>
		<td align="left"><%= oitem.FItemList(i).Fmakerid %></td>
		<td align="left"><% =oitem.FItemList(i).Fitemname %></td>
		<td align="right">
		<%
			'Response.Write "<a href=""javascript:editItemPriceInfo('" & oitem.FItemList(i).Fitemid & "')"" title='�ǸŰ� �� ���ް� ����'>" & FormatNumber(oitem.FItemList(i).Forgprice,0) & "</a>"
			Response.Write "" & FormatNumber(oitem.FItemList(i).Forgprice,0) & ""
			'���ΰ�
			if oitem.FItemList(i).Fsailyn="Y" then
				Response.Write "<br><font color=#F08050>(��)" & FormatNumber(oitem.FItemList(i).Fsailprice,0) & "</font>"
			end if
			'������
			if oitem.FItemList(i).FitemCouponYn="Y" then
				Select Case oitem.FItemList(i).FitemCouponType
					Case "1"
						'Response.Write "<br><font color=#5080F0>(��)" & FormatNumber(oitem.FItemList(i).Forgprice*((100-oitem.FItemList(i).FitemCouponValue)/100),0) & "</font>"
					Case "2"
						'Response.Write "<br><font color=#5080F0>(��)" & FormatNumber(oitem.FItemList(i).Forgprice-oitem.FItemList(i).FitemCouponValue,0) & "</font>"
				end Select
			end if
		%>
		</td>
		<td align="center"><%= fnColor(oitem.FItemList(i).Fmwdiv,"mw") %></td>
		<td align="center"><%= fnColor(oitem.FItemList(i).Fsellyn,"yn") %></td>
		<td align="center"><%= fnColor(oitem.FItemList(i).Fisusing,"yn") %></td>
		<td align="center"><%= fnColor(oitem.FItemList(i).FdeliverOverseas,"yn") %></td>
		<td align="center"><%= FormatNumber(oitem.FItemList(i).FitemWeight,0) %>g</td>
	    <td align="center">
	    	<% If oitem.FItemList(i).FExistMultiLang = "Y" Then %>
	    		���
	    	<% Else %>
	    		�̵��
	    	<% End If %>
	    </td>
	</tr>
	<% next %>
	
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15" align="center">
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
	</tr>
	
</table>
<% end if %>

<% set oitem = nothing %>

</body>
</html>
<!-- ǥ �ϴܹ� ��-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
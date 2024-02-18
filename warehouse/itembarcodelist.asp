<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/items/itemcls_2008.asp"-->
<%

dim itemid, itemname, makerid, sellyn, usingyn, danjongyn, mwdiv
dim cdl, cdm, cds
dim page
dim onoffgubun, nobarcode, noupchebarcode

itemid      = requestCheckvar(request("itemid"),255)
itemname    = request("itemname")
makerid     = requestCheckvar(request("makerid"),32)
sellyn      = requestCheckvar(request("sellyn"),10)
usingyn     = requestCheckvar(request("usingyn"),10)
danjongyn   = requestCheckvar(request("danjongyn"),10)
mwdiv       = requestCheckvar(request("mwdiv"),10)

cdl = requestCheckvar(request("cdl"),10)
cdm = requestCheckvar(request("cdm"),10)
cds = requestCheckvar(request("cds"),10)

onoffgubun = requestCheckvar(request("onoffgubun"),32)
nobarcode = requestCheckvar(request("nobarcode"),32)
noupchebarcode = requestCheckvar(request("noupchebarcode"),32)

page = requestCheckvar(request("page"),10)

if (page="") then page=1
if onoffgubun="" then onoffgubun="on"

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

set oitem = new CItem

oitem.FPageSize         = 30
oitem.FCurrPage         = page
oitem.FRectMakerid      = makerid
oitem.FRectItemid       = itemid
oitem.FRectItemName     = itemname

oitem.FRectSellYN       = sellyn
oitem.FRectIsUsing      = usingyn
oitem.FRectDanjongyn    = danjongyn
oitem.FRectMWDiv        = mwdiv

oitem.FRectCate_Large   = cdl
oitem.FRectCate_Mid     = cdm
oitem.FRectCate_Small   = cds

oitem.FRectNoBarcode   		= nobarcode
oitem.FRectNoUpcheBarcode   = noupchebarcode

if (onoffgubun = "on") then
        oitem.GetItemListByOnlineBrand
elseif (Left(onoffgubun,3) = "off") then
        oitem.FRectItemGubun =  Mid(onoffgubun,4,2)
        oitem.GetItemListByOfflineBrand
end if

dim i

%>
<script>
function NextPage(ipage){
	document.frm.page.value= ipage;
	document.frm.submit();
}

// �����Ȳ �˾�
function PopItemStock(itemid){
	var popwin = window.open("/admin/stock/itemcurrentstock.asp?menupos=709&itemid=" + itemid,"popitemstocklist","width=1000 height=600 scrollbars=yes resizable=yes");
	popwin.focus();
}

//���ڵ����
function barcodeManage(itemcode)
{
	var popbarcodemanage = window.open('/admin/stock/popBarcodeManage.asp?itemcode=' + itemcode,'popbarcodemanage','width=550,height=400,resizable=yes,scrollbars=yes');
	popbarcodemanage.focus();
}

//���ڵ����
function upcheManageCode(itemcode)
{
	var popupcheManageCode = window.open('/admin/stock/popUpcheManageCode.asp?itemcode=' + itemcode,'popupcheManageCode','width=550,height=400,resizable=yes,scrollbars=yes');
	popupcheManageCode.focus();
}

</script>

<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method=get>
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="page" >
	<tr align="center" bgcolor="<%= adminColor("topbar") %>" >
		<td rowspan="4" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
		<td align="left">
			�귣�� :<%	drawSelectBoxDesignerWithName "makerid", makerid %>
			&nbsp;
			<!-- #include virtual="/common/module/categoryselectbox.asp"-->
		</td>

		<td rowspan="4" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="�˻�" onClick="javascript:document.frm.submit();">
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("topbar") %>" >
		<td align="left">
			��ǰ�ڵ�(���ڵ�) :
			<input type="text" class="text" name="itemid" value="<%= itemid %>" size="30" maxlength="100" onKeyPress="if (event.keyCode == 13) document.frm.submit();">(��ǥ�� �����Է°���)
			&nbsp;
			��ǰ�� :
			<input type="text" class="text" name="itemname" value="<%= itemname %>" size="32" maxlength="32">
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("topbar") %>" >
		<td align="left">
			<select name="onoffgubun" >
			<option value="on" <%= ChkIIF(onoffgubun="on","selected","") %> >ON��ǰ
			<option value="off" <%= ChkIIF(onoffgubun="off","selected","") %> >OFF��ǰ
			<option value="off70" <%= ChkIIF(onoffgubun="off70","selected","") %> >OFF��ǰ-70
			<option value="off80" <%= ChkIIF(onoffgubun="off80","selected","") %> >OFF��ǰ-80
			<option value="off90" <%= ChkIIF(onoffgubun="off90","selected","") %> >OFF��ǰ-90
			</select>
			&nbsp;
			�Ǹ�:<% drawSelectBoxSellYN "sellyn", sellyn %>
	     	&nbsp;
	     	���:<% drawSelectBoxUsingYN "usingyn", usingyn %>
	     	&nbsp;
	     	����:<% drawSelectBoxDanjongYN "danjongyn", danjongyn %>
	     	&nbsp;
	     	�ŷ�����:<% drawSelectBoxMWU "mwdiv", mwdiv %>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("topbar") %>" >
		<td align="left">
			<input type="checkbox" name="nobarcode" value="Y" <%= ChkIIF(nobarcode="Y","checked","") %> > ������ڵ� ������ǰ��
			<input type="checkbox" name="noupchebarcode" value="Y" <%= ChkIIF(noupchebarcode="Y","checked","") %> > ��ü�ڵ� ������ǰ��
		</td>
	</tr>
    </form>
</table>

<p>

<!-- ����Ʈ ���� -->

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15">
			�˻���� : <b><%= oitem.FTotalCount%></b>
			&nbsp;
			������ : <b><%= page %> /<%=  oitem.FTotalpage %></b>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="90">��ǰ�ڵ�</td>
		<td width=50> �̹���</td>
		<td width="100">�귣��ID</td>
		<td>��ǰ��<br><font color="blue">[�ɼǸ�]</font></td>
		<td width="60">�ǸŰ�</td>
		<td width="60">�ɼǰ�</td>
		<td width="90">������ڵ�</td>
		<td width="90">��ü�ڵ�</td>
		<td width="225">�Է�</td>
    </tr>
<% if oitem.FresultCount<1 then %>
    <tr bgcolor="#FFFFFF">
    	<td colspan="16" align="center">[�˻������ �����ϴ�.]</td>
    </tr>
<% end if %>
<% if oitem.FresultCount > 0 then %>
    <% for i=0 to oitem.FresultCount-1 %>
	<tr class="a" height="25" bgcolor="#FFFFFF">
		<td align="center">
			<a href="http://www.10x10.co.kr/shopping/category_prd.asp?itemid=<%= oitem.FItemList(i).Fitemid %>" target="_blank" title="�̸�����">
				<% if oitem.FItemList(i).Fitemid>=1000000 then %>
			    <%= oitem.FItemList(i).Fitemgubun %>-<%= Format00(8, oitem.FItemList(i).Fitemid) %>-<%= oitem.FItemList(i).Fitemoption %>
			    <% else %>
				<%= oitem.FItemList(i).Fitemgubun %>-<%= Format00(6, oitem.FItemList(i).Fitemid) %>-<%= oitem.FItemList(i).Fitemoption %>
				<% end if %>
			</a>
		</td>
		<td align="center"><img src="<%= oitem.FItemList(i).FSmallImage %>" width="50" height="50" border="0"></td>
		<td align="left"><%= oitem.FItemList(i).Fmakerid %></td>
		<td align="left">
			<a href="javascript:PopItemStock('<% =oitem.FItemList(i).Fitemid %>')"><% =oitem.FItemList(i).Fitemname %></a><br><font color="blue">[<% =oitem.FItemList(i).Fitemoptionname %>]</font>
		</td>
		<td align="right">
			<%
				Response.Write FormatNumber(oitem.FItemList(i).Forgprice,0)
			%>
		</td>
		<td align="right">
			<%
				Response.Write FormatNumber(oitem.FItemList(i).Foptaddprice,0)
			%>
		</td>
		<td align="center"><%= oitem.FItemList(i).Fbarcode %></td>
		<td align="center"><%= oitem.FItemList(i).Fupchebarcode %></td>
		<td align="center">
			<input type="button" class="button" value="���ڵ����" onClick="barcodeManage('<%= oitem.FItemList(i).Fitemgubun %><%= CHKIIF(oitem.FItemList(i).Fitemid>=1000000,Format00(8, oitem.FItemList(i).Fitemid),Format00(6, oitem.FItemList(i).Fitemid)) %><%= oitem.FItemList(i).Fitemoption %>');">
			<input type="button" class="button" value="��ü�ڵ����" onClick="upcheManageCode('<%= oitem.FItemList(i).Fitemgubun %><%= CHKIIF(oitem.FItemList(i).Fitemid>=1000000,Format00(8, oitem.FItemList(i).Fitemid),Format00(6, oitem.FItemList(i).Fitemid)) %><%= oitem.FItemList(i).Fitemoption %>');">
		</td>
	</tr>
	<% next %>

	<tr height="25" bgcolor="FFFFFF">
		<td colspan="16" align="center">
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


<!-- ǥ �ϴܹ� ��-->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
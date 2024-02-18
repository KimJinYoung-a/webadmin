<%@ language=vbscript %>
<% option explicit %>
<%
'###############################################
' Discription : ����� mdpick
' History : 2013.12.17 �ѿ��
'###############################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/offshop_function.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/mobile/main/inc_mainhead.asp"-->
<!-- #include virtual="/lib/classes/mobile/mdpick_cls.asp" -->
<%
Dim isusing, page, i, okeyword, reload, itemid, itemname, makerid, sellyn, usingyn, danjongyn
dim mwdiv, limityn, vatyn, sailyn, overSeaYn, itemdiv, cdl, cdm, cds, dispCate, acURL
	page = request("page")
	reload = request("reload")
	isusing = RequestCheckVar(request("isusing"),1)
	itemid      = requestCheckvar(request("itemid"),255)
	itemname    = request("itemname")
	makerid     = requestCheckvar(request("makerid"),32)
	sellyn      = requestCheckvar(request("sellyn"),10)
	usingyn     = requestCheckvar(request("usingyn"),10)
	danjongyn   = requestCheckvar(request("danjongyn"),10)
	mwdiv       = requestCheckvar(request("mwdiv"),10)
	limityn     = requestCheckvar(request("limityn"),10)
	vatyn       = requestCheckvar(request("vatyn"),10)
	sailyn      = requestCheckvar(request("sailyn"),10)
	overSeaYn   = requestCheckvar(request("overSeaYn"),10)
	itemdiv     = requestCheckvar(request("itemdiv"),10)
	cdl = requestCheckvar(request("cdl"),10)
	cdm = requestCheckvar(request("cdm"),10)
	cds = requestCheckvar(request("cds"),10)
	dispCate = requestCheckvar(request("disp"),16)

acURL =Server.HTMLEncode("/admin/mobile/mdpick/mdpick_process.asp?menupos="&menupos)
	
if itemid<>"" then
	dim iA ,arrTemp,arrItemid

	arrTemp = Split(itemid,chr(13))

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

if page="" then page=1
if reload="" and isusing="" then isusing="Y"

set okeyword = new cmdpick
	okeyword.FPageSize		= 100
	okeyword.FCurrPage		= page
	okeyword.Frectisusing			= isusing
	okeyword.FRectMakerid      = makerid
	okeyword.FRectItemid       = itemid
	okeyword.FRectItemName     = itemname
	okeyword.FRectSellYN       = sellyn
	okeyword.FRectitemIsUsing      = usingyn
	okeyword.FRectDanjongyn    = danjongyn
	okeyword.FRectLimityn      = limityn
	okeyword.FRectMWDiv        = mwdiv
	okeyword.FRectVatYn        = vatyn
	okeyword.FRectSailYn       = sailyn
	okeyword.FRectIsOversea	= overSeaYn
	okeyword.FRectCate_Large   = cdl
	okeyword.FRectCate_Mid     = cdm
	okeyword.FRectCate_Small   = cds
	okeyword.FRectDispCate		= dispCate
	okeyword.FRectItemDiv      = itemdiv
	okeyword.getmdpick_list()

%>
<link href="/js/jqueryui/css/jquery-ui.css" rel="stylesheet">
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/jqueryui/jquery-ui-1.10.2.custom.min.js"></script>
<script type='text/javascript'>

function totalCheck(){
	var f = document.frmlist;
	var objStr = "idx";
	var chk_flag = true;
	for(var i=0; i<f.elements.length; i++) {
		if(f.elements[i].name == objStr) {
			if(!f.elements[i].checked) {
				chk_flag = f.elements[i].checked;
				break;
			}
		}
	}

	for(var i=0; i < f.elements.length; i++) {
		if(f.elements[i].name == objStr) {
			if(chk_flag) {
				f.elements[i].checked = false;
			} else {
				f.elements[i].checked = true;
			}
		}
	}
}

function frmsubmit(page){
	frm.page.value=page;
	frm.submit();
}

// ��ǰ�߰�(�˻�) �˾�
function addnewItem(){
	var popwin;
	popwin = window.open("/admin/itemmaster/pop_itemAddInfo.asp?acURL=<%=acURL%>&menupos=<%=menupos%>", "popup_item", "width=1024,height=768,scrollbars=yes,resizable=yes");
	popwin.focus();
}

function mdpickedit(idx){
	var mdpickedit = window.open('/admin/mobile/mdpick/mdpick_edit.asp?idx='+idx+'&menupos=<%=menupos%>','mdpickedit','width=1024,height=768,scrollbars=yes,resizable=yes');
	mdpickedit.focus();
}

function AssignXmlReal(){
	if (confirm('����ϻ���Ʈ ���� �������� ���� �Ͻðڽ��ϱ�?')){
		 var popwin = window.open('','refreshFrm','');
		 popwin.focus();
		 refreshFrm.target = "refreshFrm";
		 refreshFrm.action = "<%=mobileUrl%>/chtml/mobile/make_mdpick_xml.asp" ;
		 refreshFrm.submit();
	}
}

//�ּ�ó��
function AssignXmlAppl(term){
    if (!confirm('���� �ݿ��Ͻðڽ��ϱ�?')) return;
     
	 var popwin = window.open('','refreshFrm_Main','');
	 popwin.focus();
	 refreshFrm.target = "refreshFrm_Main";
	 refreshFrm.action = "<%=mobileUrl%>/chtml/mobile/make_mdpick_xml.asp?term=" + term;
	 refreshFrm.submit();
}

</script>

<img src="/images/icon_arrow_link.gif"> <b>MDPICK</b>
<p>
<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="page" value="1">
<input type="hidden" name="reload" value="ON">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
		* MDPICK ��뿩�� : <% DrawSelectBoxUsingYN "isusing",isusing %>
		<p>
		* �귣�� : <% drawSelectBoxDesignerWithName "makerid", makerid %>
		&nbsp;&nbsp;
		* ��ǰ�� : <input type="text" class="text" name="itemname" value="<%= itemname %>" size="32" maxlength="32">
		&nbsp;&nbsp;
		<span style="white-space:nowrap;">* ��ǰ�ڵ� :
		<input type="text" class="text" name="itemid" value="<%= itemid %>" size="30" maxlength="100" onKeyPress="if (event.keyCode == 13) document.frm.submit();">(��ǥ�� �����Է°���)</span>
		<p>
		<span style="white-space:nowrap;">* �Ǹ� : <% drawSelectBoxSellYN "sellyn", sellyn %></span>
     	&nbsp;&nbsp;
     	<span style="white-space:nowrap;">* ��ǰ��� : <% drawSelectBoxUsingYN "usingyn", usingyn %></span>
     	&nbsp;&nbsp;
     	<span style="white-space:nowrap;">* ���� : <% drawSelectBoxDanjongYN "danjongyn", danjongyn %></span>
     	&nbsp;&nbsp;
     	<span style="white-space:nowrap;">* ���� : <% drawSelectBoxLimitYN "limityn", limityn %></span>
     	&nbsp;&nbsp;
     	<span style="white-space:nowrap;">* �ŷ����� : <% drawSelectBoxMWU "mwdiv", mwdiv %></span>
     	&nbsp;&nbsp;
     	<span style="white-space:nowrap;">* ���� : <% drawSelectBoxVatYN "vatyn", vatyn %></span>
     	&nbsp;&nbsp;
     	<span style="white-space:nowrap;">* ���� : <% drawSelectBoxSailYN "sailyn", sailyn %></span>
     	&nbsp;&nbsp;
     	<span style="white-space:nowrap;">* �ؿܹ�� : <% drawSelectBoxIsOverSeaYN "overSeaYn", overSeaYn %></span>
        &nbsp;&nbsp;
     	<span style="white-space:nowrap;">* ��ǰ���� : <% drawSelectBoxItemDiv "itemdiv", itemdiv %></span>
		<p>
		* ����<!-- #include virtual="/common/module/categoryselectbox.asp"-->
		&nbsp;&nbsp;����ī�װ� : <!-- #include virtual="/common/module/dispCateSelectBox.asp"-->
	</td>
	<td width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="javascript:frmsubmit('');">
	</td>
</tr>
</form>
<form name="refreshFrm" method="post">
</form>
</table>
<!-- �˻� �� -->

<br>

<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding:10 0 10 0;">
<tr>
	<td>
		<a href="javascript:AssignXmlReal();"><img src="/images/refreshcpage.gif" border="0"> Real ����</a>
		<!--������ �����Ͽ� <input type="text" name="vTerm" value="1" size="1" class="text" style="text-align:right;">�ϰ�
		<a href="javascript:AssignXmlAppl(document.all.vTerm.value);"><img src="/images/refreshcpage.gif" border="0"> XML Real ����(����)</a>-->
	</td>
    <td align="right">
    	<input type="button" value="��ǰ�߰�(�˻�)" onclick="addnewItem();" class="button">
    </td>
</tr>
</table>

<!--  ����Ʈ -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="20">
		�� ��ϼ� : <b><%=okeyword.FtotalCount%></b>
		&nbsp;
		������ : <b><%= page %> / <%=okeyword.FtotalPage%></b>
	</td>
</tr>
<tr height="25" align="center" bgcolor="<%= adminColor("tabletop") %>">
	<!--<td><input type="checkbox" name="ckall" onclick="totalCheck()"></td>-->
	<td>�̹���</td>
	<td>��ǰ�ڵ�</td>	
    <td>��ǰ��</td>
    <!--<td>������</td>
    <td>������</td>-->
    <td>���ļ���</td>
    <td>��뿩��</td>
    <td>���</td>
</tr>
<form name="frmlist" method="post">
<%
if okeyword.FResultCount>0 then
	
for i=0 to okeyword.FResultCount - 1 
%>
<tr height="30" align="center" bgcolor="<%=chkIIF(okeyword.FItemList(i).fisusing="Y","#FFFFFF","#F0F0F0")%>">
	<!--<td><input type="checkbox" name="idx" value="<%=okeyword.FItemList(i).Fidx%>" onClick="AnCheckClick(this);"></td>-->
    <td>
    	<img src="<%= okeyword.FItemList(i).fbasicimage %>" width=50 height=50 />
	</td>
	<td>
		<%= okeyword.FItemList(i).fitemid %>
	</td>	
	<td>
		<%= okeyword.FItemList(i).Fitemname %>
	</td>
    <!--<td>
    	<% if okeyword.FItemList(i).FStartdate<>"" then %>
    		<%= okeyword.FItemList(i).FStartdate %>
    	<% end if %>
    </td>
    <td>
    	<% if okeyword.FItemList(i).FStartdate<>"" then %>
		    <% if (okeyword.FItemList(i).IsEndDateExpired) then %>
		    	<font color="#777777"><%= Left(okeyword.FItemList(i).FEnddate,10) %></font>
		    <% else %>
		    	<%= Left(okeyword.FItemList(i).FEnddate,10) %>
		    <% end if %>
    	<% end if %>		    
    </td>-->	
	<td>
		<%= okeyword.FItemList(i).forderno %>
	</td>
	<td><%= okeyword.FItemList(i).fisusing %></td>
	<td>
		<input type="button" onclick="mdpickedit('<%=okeyword.FItemList(i).Fidx%>')" value="����" class="button">
	</td>
</tr>
<% Next %>

<tr bgcolor="#FFFFFF">
	<td align="center" colspan="20">
		<% if okeyword.HasPreScroll then %>
			<span class="list_link"><a href="javascript:frmsubmit('<%= okeyword.StartScrollPage-1 %>')">[pre]</a></span>
		<% else %>
		[pre]
		<% end if %>
		<% for i = 0 + okeyword.StartScrollPage to okeyword.StartScrollPage + okeyword.FScrollCount - 1 %>
			<% if (i > okeyword.FTotalpage) then Exit for %>
			<% if CStr(i) = CStr(okeyword.FCurrPage) then %>
			<span class="page_link"><font color="red"><b><%= i %></b></font></span>
			<% else %>
			<a href="javascript:frmsubmit('<%= i %>')" class="list_link"><font color="#000000"><%= i %></font></a>
			<% end if %>
		<% next %>
		<% if okeyword.HasNextScroll then %>
			<span class="list_link"><a href="javascript:frmsubmit('<%= i %>')">[next]</a></span>
		<% else %>
		[next]
		<% end if %>
	</td>
</tr>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="20" align="center" class="page_link">[�˻������ �����ϴ�.]</td>
	</tr>
<% end if %>
</form>
</table>

<%
set okeyword = Nothing
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
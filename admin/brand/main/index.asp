<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  �귣�彺Ʈ��Ʈ
' History : 2013.08.19 ������ ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/street/brandmainCls.asp" -->
<%
Dim page, lbrand, i
Dim chgMode
page    = request("page")
chgMode = request("chgMode")

If page = "" Then page = 1

'// ��� ����
Set lbrand = New cBrandMain
	lbrand.FCurrPage = page
	lbrand.FPageSize=20
	lbrand.FRectGubun=1
	lbrand.sMainTop3List
%>

<script language="javascript">

function AnSelectAllFrame(bool){
	var frm;
	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			if (frm.cksel.disabled!=true){
				frm.cksel.checked = bool;
				AnCheckClick(frm.cksel);
			}
		}
	}
}

function AnCheckClick(e){
	if (e.checked)
		hL(e);
	else
		dL(e);
}	

function ckAll(icomp){
	var bool = icomp.checked;
	AnSelectAllFrame(bool);
}

function CheckSelected(){
	var pass=false;
	var frm;

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			pass = ((pass)||(frm.cksel.checked));
		}
	}

	if (!pass) {
		return false;
	}
	return true;
}

function AssignXmlReal(upfrm,imagecount){
	if (!CheckSelected()){
		alert('���þ������� �����ϴ�.');
		return;
	}
	var frm;

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			if (frm.cksel.checked){
				upfrm.fidx.value = upfrm.fidx.value + frm.idx.value + "," ;
			}
		}
	}
	var tot;
	tot = upfrm.fidx.value;
	upfrm.fidx.value = ""

	var AssignimageReal;
	AssignimageReal = window.open("", "AssignimageReal","width=800,height=600,scrollbars=yes,resizable=yes");
	AssignimageReal.location.href="<%=wwwUrl%>/chtml/street/Main_Top3BannerJS.asp?idx=" +tot + '&imagecount='+imagecount;
	AssignimageReal.focus();
}

//�̹����űԵ�� & ����
function AddNewMainContents(idx){
	var AddNewMainContents = window.open('/admin/brand/main/imagemake_contents.asp?idx='+ idx + '&gubun=1','AddNewMainContents','width=1200,height=600,scrollbars=yes,resizable=yes');
	AddNewMainContents.focus();
}

//���� ����
function jsSort() {
	if (!CheckSelected()){
		alert('���þ������� �����ϴ�.');
		return;
	}
	var frm;

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			if (frm.cksel.checked){
				document.frm.fidx.value = document.frm.fidx.value + frm.idx.value + "," ;
				document.frm.sortnoarr.value = document.frm.sortnoarr.value  + frm.sortNo.value + ",";
			}
		}
	}
	document.frm.mode.value = '3banner';
	document.frm.action = '/admin/brand/main/mainSortnoProcess.asp';
	document.frm.submit();
}

function chgMAINREG(val){
	if(val == "1"){
		location.replace('/admin/brand/main/index.asp?menupos=<%=menupos%>');
	}else if(val == "2"){
		location.replace('/admin/brand/main/brandPick.asp?chgMode=2&menupos=<%=menupos%>');
	}else if(val == "3"){
		location.replace('/admin/brand/main/mainInterView.asp?chgMode=3&menupos=<%=menupos%>');
	}else if(val == "4"){
		location.replace('/admin/brand/main/mainLookBook.asp?chgMode=4&menupos=<%=menupos%>');
	}else if(val == "5"){
		window.open('<%=wwwUrl%>/chtml/street/taglist.asp','','width=450,height=130,scrollbars=no');
	}
}

document.domain ='10x10.co.kr';

</script>

<!-- #include virtual="/admin/brand/inc_streetHead.asp"-->

<img src="/images/icon_arrow_link.gif"> <b>��������������</b>
<!-- �˻� ���� -->
<form name="frm" method="get" action="" style="margin:0px;">
<input type="hidden" name="page" value="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="mode">
<input type="hidden" name="fidx">
<input type="hidden" name="sortnoarr" value="">
<!--<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
	</td>	
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="frm.submit()">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
	</td>
</tr>
</table>-->
<!-- �˻� �� -->

<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
		<select name="chgMode" class="select" onchange="javascrtip:chgMAINREG(this.value);">
			<option value="1">����TOP3 �Ѹ����</option>
			<option value="2" <%= chkIIF(chgMode="2","selected","") %>>����BRAND PICK</option>
			<option value="3" <%= chkIIF(chgMode="3","selected","") %>>����InterView</option>
			<option value="4" <%= chkIIF(chgMode="4","selected","") %>>����LookBook</option>
			<option value="5" <%= chkIIF(chgMode="5","selected","") %>>����BRAND TAG</option>
		</select>
		&nbsp;&nbsp;&nbsp;&nbsp;
		<a href="javascript:AssignXmlReal(frm,3);"><img src="/images/refreshcpage.gif" border="0"> Real ����</a>
		&nbsp;&nbsp;&nbsp;&nbsp;
		<input class="button" type="button" id="btnEditSel" value="�켱��������" onClick="jsSort();">
	</td>
	<td align="right">
		<input type="button" value="�űԵ��" class="button" onClick="javascript:AddNewMainContents('0');">
	</td>
</tr>
</table>
<!-- �׼� �� -->

</form>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		�˻���� : <b><%=lbrand.FTotalCount %></b>
		&nbsp;
		������ : <b><%= page %>/ <%= lbrand.FTotalPage %></b>	
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td align="center"><input type="checkbox" name="ckall" onclick="ckAll(this)"></td>
    <td align="center">Image</td>
    <td align="center">�켱����</td>
    <td align="center">�����</td>
</tr>
<% If lbrand.FResultCount > 0 Then %> 
<% For i = 0 to lbrand.FResultCount - 1 %>
<tr align="center" <%= chkiif(lbrand.FItemList(i).FIsusing="N", "bgcolor='#DDDDDD'", "bgcolor='#FFFFFF'") %> >
<form action="" name="frmBuyPrc<%=i%>" method="get">
<input type="hidden" name="idx" value="<%= lbrand.FItemList(i).Fidx %>">
	<td align="center"><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"></td>
    <td align="center">
    	<a href="javascript:AddNewMainContents('<%= lbrand.FItemlist(i).Fidx %>');">
    	<img width=40 height=40 src="<%=uploadUrl%>/brandstreet/main/<%= lbrand.FItemlist(i).FImagepath %>" border="0">
    	</a>
    </td>
    <td align="center"><input type="text" size="2" maxlength="2" name="sortNo" value="<%=lbrand.FItemList(i).FImage_order%>" class="text"></td>
    <td align="center"><%= lbrand.FItemlist(i).FRegdate %></td>
</tr>
</form>
<% Next %>
<% Else %>
<tr bgcolor="#FFFFFF" height="30">
	<td colspan="5" align="center" class="page_link">[��ϵ� �����Ͱ� �����ϴ�.]</td>
</tr>
<% End If %>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15" align="center">
       	<% If lbrand.HasPreScroll Then %>
			<span class="list_link"><a href="?page=<%= lbrand.StartScrollPage-1 %>">[pre]</a></span>
		<% Else %>
		[pre]
		<% End If %>
		<% for i = 0 + lbrand.StartScrollPage to lbrand.StartScrollPage + lbrand.FScrollCount - 1 %>
			<% if (i > lbrand.FTotalpage) then Exit for %>
			<% if CStr(i) = CStr(lbrand.FCurrPage) then %>
			<span class="page_link"><font color="red"><b><%= i %></b></font></span>
			<% else %>
			<a href="?page=<%= i %>" class="list_link"><font color="#000000"><%= i %></font></a>
			<% end if %>
		<% next %>
		<% if lbrand.HasNextScroll then %>
			<span class="list_link"><a href="?page=<%= i %>">[next]</a></span>
		<% else %>
		[next]
		<% end if %>
	</td>
</tr>
</table>

<%
set lbrand = nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
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
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/street/lookbookCls.asp"-->
<%
Dim idx, olist, page, i, state, title, makerid, isusing, research
	page	= request("page")
	idx		= request("idx")
	state	= request("state")
	title	= request("title")
	makerid	= request("makerid")
	isusing	= request("isusing")
	research	= request("research")
	menupos	= request("menupos")

Dim chgMode
chgMode = request("chgMode")

If page = ""	Then page = 1
if research ="" and state="" then state = "7"
If isusing = "" Then isusing = "Y"

SET olist = new clookbook
	olist.FCurrPage		= page
	olist.FPageSize		= 20
	olist.FrectMakerid		= makerid
	olist.Frectstate		= state
	olist.Frecttitle		= title
	olist.frectisusing = isusing
	olist.sblookBookmasterAdminlist
%>

<script language="javascript">

var ichk;
ichk = 1;

function jsChkAll(){
	var frm, blnChk;
	frm = document.fitem;
	if(!frm.chkI) return;
	if ( ichk == 1 ){
		blnChk = true;
		ichk = 0;
	}else{
		blnChk = false;
		ichk = 1;
	}
	for (var i=0;i<frm.elements.length;i++){
		var e = frm.elements[i];
		if ((e.type=="checkbox")) {
			e.checked = blnChk ;
		}
	}
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
// �̹��� Ŭ���� ���� ũ��� �˾� ����
function doImgPop(img){
	img1= new Image();
	img1.src=(img);
	imgControll(img);
}

function imgControll(img){
	if((img1.width!=0)&&(img1.height!=0)){
		viewImage(img);
	}else{
		controller="imgControll('"+img+"')";
		intervalID=setTimeout(controller,20);
	}
}

function viewImage(img){
	W=img1.width;
	H=img1.height;
	O="width="+W+",height="+H+",scrollbars=yes";
	imgWin=window.open("","",O);
	imgWin.document.write("<html><head><title>:*:*:*: �̹����󼼺��� :*:*:*:*:*:*:</title></head>");
	imgWin.document.write("<body topmargin=0 leftmargin=0>");
	imgWin.document.write("<img src="+img+" onclick='self.close()' style='cursor:pointer;' title ='Ŭ���Ͻø� â�� �����ϴ�.'>");
	imgWin.document.close();
}

function gosubmit(page){
    var frm = document.frm;
    frm.page.value=page;
	frm.submit();
}
function CheckSelected(){
	var pass=false;
	var frm;

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			pass = ((pass)||(frm.chkI.checked));
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
			if (frm.chkI.checked){
				upfrm.fidx.value = upfrm.fidx.value + frm.idx.value + "," ;
				if (!frm.sortNo.value){
					frm.sortNo.value = "0";
				}
				upfrm.sortnoarr.value = upfrm.sortnoarr.value + frm.sortNo.value + "," ;
			}
		}
	}
	var tot;
	var tot2;
	tot = upfrm.fidx.value;
	tot2 = upfrm.sortnoarr.value;
	upfrm.fidx.value = ""
	upfrm.sortnoarr.value = ""


	var AssignimageReal;
	AssignimageReal = window.open("", "AssignimageReal","width=800,height=600,scrollbars=yes,resizable=yes");
	AssignimageReal.location.href="<%=wwwUrl%>/chtml/street/Main_LookBookJS.asp?idx=" +tot + '&sort='+tot2+'&imagecount='+imagecount;
	AssignimageReal.focus();
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
			if (frm.chkI.checked){
				document.frm.fidx.value = document.frm.fidx.value + frm.idx.value + "," ;
				document.frm.sortnoarr.value = document.frm.sortnoarr.value  + frm.sortNo.value + ",";
			}
		}
	}
	document.frm.mode.value = 'lookbook';
	document.frm.action = '/admin/brand/main/mainSortnoProcess.asp';
	document.frm.submit();
}
</script>
<!-- #include virtual="/admin/brand/inc_streetHead.asp"-->

<img src="/images/icon_arrow_link.gif"> <b>��������������</b>

<!-- �˻� ���� -->
<form name="frm" method="get" action="" style="margin:0px;">
<input type="hidden" name="page" value="<%=page%>">
<input type="hidden" name="research" value="on">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="fidx">
<input type="hidden" name="sortnoarr">
<input type="hidden" name="mode">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
		* �귣�� : <% LookBook_ID_with_Name "makerid",makerid , " onchange='gosubmit("""");'" %>
		&nbsp;&nbsp;
		���� : <input type="text" name="title" value="<%=title%>" size="40" maxlength="40" class="text">
		&nbsp;&nbsp;
		* ���� :
		<% drawlookbookstats "state" , state , " onchange='gosubmit("""");'" %>
		&nbsp;&nbsp;
		* ������� :
		<% drawSelectBoxUsingYN "isusing", isusing %>
	</td>	
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="gosubmit('');">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
	</td>
</tr>
</table>
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
		<input class="button" type="button" id="btnEditSel" value="���Ĺ�ȣ����" onClick="jsSort();">
	</td>
	<td align="right">
	</td>
</tr>
</table>
</form>
<!-- �׼� �� -->

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<input type="hidden" name="sortnoarr" value="">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="20">
		�˻���� : <b><%=olist.FTotalCount %></b>
		&nbsp;
		������ : <b><%= page %>/ <%= olist.FTotalPage %></b>		
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="30">
	<td></td>
	<td>��ȣ</td>
	<td>�귣��</td>
	<td>�̹���</td>
	<td>����</td>
	<td>����</td>
	<td>�������</td>
	<td>��뿩��</td>	
</tr>
<% if olist.fresultcount > 0 then %>
<% For i = 0 to olist.fresultcount -1 %>
<form action="" name="frmBuyPrc<%=i%>" method="get">
<input type="hidden" name="idx" value="<%= olist.FItemList(i).fidx %>">
<% if olist.FItemList(i).fisusing="Y" then %>
	<tr height="25" bgcolor="FFFFFF"  align="center">
<% else %>
	<tr height="25" bgcolor="f1f1f1"  align="center">
<% end if %>	
	<td align="center"><input type="checkbox" name="chkI" onClick="AnCheckClick(this);"  value="<%= olist.FItemList(i).FIdx %>"></td>
	<td align="center"><%= olist.FItemList(i).fidx %></td>
	<td align="center"><%= olist.FItemList(i).FMakerid %></td>
	<td align="center">
		<img src="<%=olist.FItemList(i).fmainimg%>" width="50" height="50" title="Ŭ���Ͻø� ����ũ��� ���� �� �ֽ��ϴ�." style="cursor: pointer;" onclick="doImgPop('<%=olist.FItemList(i).fmainimg%>')"/>
	</td>
	<td>
		<%= olist.FItemList(i).Ftitle %>
	</td>
	<td><input type="text" size="2" maxlength="2" name="sortNo" value="<%= olist.FItemList(i).FmainpageSortNo %>" class="text"></td>
	<td>
		<%= lookbookstatsname(olist.FItemList(i).Fstate) %>
	</td>
	<td>
		<%=olist.FItemList(i).FIsusing%>
	</td>	
</tr>
</form>
<% Next %>
<tr height="25" bgcolor="FFFFFF" >
	<td colspan="15" align="center">
       	<% If olist.HasPreScroll Then %>
			<span class="olist_link"><a href="javascript:gosubmit('<%= olist.StartScrollPage-1 %>');">[pre]</a></span>
		<% Else %>
		[pre]
		<% End If %>
		<% For i = 0 + olist.StartScrollPage to olist.StartScrollPage + olist.FScrollCount - 1 %>
			<% If (i > olist.FTotalpage) Then Exit for %>
			<% If CStr(i) = CStr(olist.FCurrPage) Then %>
			<span class="page_link"><font color="red"><b><%= i %></b></font></span>
			<% Else %>
			<a href="javascript:gosubmit('<%= i %>');" class="olist_link"><font color="#000000"><%= i %></font></a>
			<% End if %>
		<% Next %>
		<% If olist.HasNextScroll Then %>
			<span class="olist_link"><a href="javascript:gosubmit('<%= i %>');">[next]</a></span>
		<% Else %>
		[next]
		<% End If %>
	</td>
</tr>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="20" align="center" class="page_link">[�˻������ �����ϴ�.]</td>
	</tr>
<% end if %>
</table>
<% 
SET olist = nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
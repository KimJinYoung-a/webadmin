<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  ����Ʈ
' History : 2014.03.19 �ѿ�� ����
' History : 2014.10.31 ���¿� mtitle �߰�
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/sitemaster/gift/giftday_cls.asp"-->
<%
dim research, title, mtitle, isusing, page, masteridx, cgiftday, i
	title	= requestcheckvar(request("title"),128)
	mtitle	= requestcheckvar(request("mtitle"),128)
	page	= requestcheckvar(request("page"),10)
	isusing	= requestcheckvar(request("isusing"),1)
	research	= requestcheckvar(request("research"),2)
	menupos	= requestcheckvar(request("menupos"),10)
	masteridx	= requestcheckvar(request("masteridx"),10)
	
If page = ""	Then page = 1
if research ="" and isusing="" then isusing = "Y"

SET cgiftday = new Cgiftday_list
	cgiftday.FCurrPage		= page
	cgiftday.FPageSize		= 50
	cgiftday.Frecttitle		= title
	cgiftday.Frectmtitle		= mtitle
	cgiftday.Frectisusing		= isusing
	cgiftday.Frectmasteridx		= masteridx
	cgiftday.getgiftday_master
%>

<script type='text/javascript'>

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

function giftdayedit(masteridx){
	var giftdayedit = window.open('/admin/sitemaster/gift/day/giftday_edit.asp?masteridx='+masteridx+'&menupos=<%=menupos%>','giftdayedit','width=1024,height=768,scrollbars=yes,resizable=yes');
	giftdayedit.focus();
}

function giftdaywinner(masteridx){
	var giftdaywinner = window.open('/admin/sitemaster/gift/day/giftdaywinner.asp?masteridx='+masteridx+'&menupos=<%=menupos%>','giftdaywinner','width=1024,height=768,scrollbars=yes,resizable=yes');
	giftdaywinner.focus();
}

function gosubmit(page){
    var frm = document.frm;
    frm.page.value=page;
	frm.submit();
}

function jsSetItem(idx){
	var popitem;
	popitem = window.open('/admin/sitemaster/gift/day/giftday_item.asp?idx='+idx,'popitem','width=920,height=600,scrollbars=yes,resizable=yes');
	popitem.focus();
}

</script>

<!-- �˻� ���� -->
<form name="frm" method="get" action="" style="margin:0px;">
<input type="hidden" name="page" value="<%=page%>">
<input type="hidden" name="research" value="on">
<input type="hidden" name="menupos" value="<%= menupos %>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
		* ��ȣ : <input type="text" name="masteridx" value="<%=masteridx%>" size="10" maxlength="10" class="text">
		&nbsp;&nbsp;
		* ���� : <input type="text" name="title" value="<%=title%>" size="40" maxlength="40" class="text">
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
</form>
<!-- �˻� �� -->

<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
	</td>
	<td align="right">
		<input type="button" value="�����űԵ��" class="button" onclick="giftdayedit('');">
	</td>
</tr>
</table>
<!-- �׼� �� -->

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="fitem" method="post" style="margin:0px;">
<input type="hidden" name="sortnoarr" value="">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="20">
		�˻���� : <b><%=cgiftday.FTotalCount %></b>
		&nbsp;
		������ : <b><%= page %>/ <%= cgiftday.FTotalPage %></b>		
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="30">
	<!--<td><input type="checkbox" name="chkA" onClick="jsChkAll();"></td>-->
	<td>��ȣ</td>
	<td>WWW<Br>����Ʈž</td>
	<td>����</td>
	<td>���������</td>
	<td>�Ⱓ</td>
	<td>���<Br>����</td>
	<td>�翬��</td>
	<td>���</td>
</tr>
<% if cgiftday.fresultcount > 0 then %>
<% For i = 0 to cgiftday.fresultcount -1 %>
<% if cgiftday.FItemList(i).fisusing="Y" then %>
	<tr height="25" bgcolor="FFFFFF"  align="center">
<% else %>
	<tr height="25" bgcolor="f1f1f1"  align="center">
<% end if %>	
	<!--<td align="center"><input type="checkbox" name="chkI" onClick="AnCheckClick(this);"  value="<%= cgiftday.FItemList(i).fmasteridx %>"></td>-->
	<td align="center"><%= cgiftday.FItemList(i).fmasteridx %></td>
	<td align="center">
		<img src="<%=cgiftday.FItemList(i).flisttopimg_w%>" width="50" height="50" title="Ŭ���Ͻø� ����ũ��� ���� �� �ֽ��ϴ�." style="cursor: pointer;" onclick="doImgPop('<%=cgiftday.FItemList(i).flisttopimg_w%>')"/>
	</td>
	<td align="center"><%= ReplaceBracket(cgiftday.FItemList(i).ftitle) %></td>
	<td align="center"><%= ReplaceBracket(cgiftday.FItemList(i).fmtitle) %></td>
	<td align="center"><%= left(cgiftday.FItemList(i).fstartdate,10) %> - <%= left(cgiftday.FItemList(i).fenddate,10) %></td>
	<td><%=cgiftday.FItemList(i).FIsusing%></td>
	<td><%=cgiftday.FItemList(i).fdetailcount%></td>
	<td>
		<input type="button" onClick="giftdayedit('<%=cgiftday.FItemList(i).fmasteridx%>');" value="����" class="button">
		<input type="button" onClick="giftdaywinner('<%=cgiftday.FItemList(i).fmasteridx%>');" value="��������Ʈ" class="button">
		<input type="button" class="button" value="��ǰȮ��[<%= cgiftday.FItemList(i).Fitemcnt %>]" onclick="jsSetItem('<%= cgiftday.FItemList(i).fmasteridx %>','0');"/>
	</td>	
</tr>
<% Next %>
<tr height="25" bgcolor="FFFFFF" >
	<td colspan="15" align="center">
       	<% If cgiftday.HasPreScroll Then %>
			<span class="cgiftday_link"><a href="javascript:gosubmit('<%= cgiftday.StartScrollPage-1 %>');">[pre]</a></span>
		<% Else %>
		[pre]
		<% End If %>
		<% For i = 0 + cgiftday.StartScrollPage to cgiftday.StartScrollPage + cgiftday.FScrollCount - 1 %>
			<% If (i > cgiftday.FTotalpage) Then Exit for %>
			<% If CStr(i) = CStr(cgiftday.FCurrPage) Then %>
			<span class="page_link"><font color="red"><b><%= i %></b></font></span>
			<% Else %>
			<a href="javascript:gosubmit('<%= i %>');" class="cgiftday_link"><font color="#000000"><%= i %></font></a>
			<% End if %>
		<% Next %>
		<% If cgiftday.HasNextScroll Then %>
			<span class="cgiftday_link"><a href="javascript:gosubmit('<%= i %>');">[next]</a></span>
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
</form>
</table>

<% 
SET cgiftday = nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
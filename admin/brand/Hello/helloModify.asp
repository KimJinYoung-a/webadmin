<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  �귣�彺Ʈ��Ʈ
' History : 2013.08.29 ������ ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/street/helloCls.asp"-->
<!-- #include virtual="/lib/classes/street/managerCls.asp"-->
<%
Dim ohello, makerid, mode
Dim omanager, menugubun
makerid = request("makerid")
mode = request("mode")

SET ohello = new chello
	ohello.FRectMakerid	= makerid
	ohello.sbhellomodify

If ohello.FTotalCount <> 0 Then
	mode = "U"
	SET omanager = new cmanager
		omanager.FRectMakerid	= makerid
		omanager.sbbrandgubunlist_confirm
		If omanager.Ftotalcount > 0 Then
			menugubun = omanager.FOneItem.fbrandgubun
		End If
	SET omanager = nothing
End If
%>
<script language="javascript">
function form_check(){
	var frm = document.frmetc;
	if(frm.makerid.value==""){
		alert('�귣�带 �Է��ϼ���');
		frm.makerid.focus();
		return false;
	}

	if (confirm('�귣��Hello�κ��� ���� �Ͻðڽ��ϱ�?')){
		frm.submit();
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
function duplProcess(){
	var strid;
	strid = document.frmetc.makerid.value;
	if(strid == ""){
		alert('�귣�带 �Է��ϼ���');
		document.frmetc.makerid.focus();
		return;	
	}
    document.dupl.target = "xLink";
    document.dupl.duplid.value = document.frmetc.makerid.value;
    document.dupl.action = "duplprocess.asp"
    document.dupl.submit();
}
function imgDelProcess(){
	if (confirm('�̹����� ���� �Ͻðڽ��ϱ�?')){
	    document.dupl.target = "xLink";
	    document.dupl.duplid.value = "<%=makerid%>";
	    document.dupl.action = "imgDelprocess.asp"
	    document.dupl.submit();
	}
}
</script>
<!-- #include virtual="/admin/brand/inc_streetHead.asp"-->
<form name="dupl">
	<input type="hidden" name="duplid" value="">
</form>
<table border="0" cellpadding="0" cellspacing="0" class="a">
<tr height="30"><td><img src="/images/icon_arrow_link.gif"></td><td style="padding-top:3">&nbsp;<b>Hello �귣�� �Ұ�</b></td></tr>
</table>
<table width="100%", cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmetc" method="post" action="<%= uploadUrl %>/linkweb/street/doHellouploadAdm.asp" enctype="MULTIPART/FORM-DATA">
<input type="hidden" name="mode" value="<%=mode%>">
<% If ohello.FTotalCount = 0 Then %>
<tr>
	<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">�귣��</td>
	<td bgcolor="#FFFFFF" colspan="3"><% drawSelectBoxDesignerwithName "makerid",makerid %>
		&nbsp;<span onclick="duplProcess();" style="cursor:pointer;">�ߺ�Ȯ��</span>
	</td>
</tr>
<% Else %>
	<input type= "hidden" name="makerid" value="<%=makerid%>">
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>" >�귣���(�ѱ�)</td>
	<td bgcolor="FFFFFF"><input type="text" value="<%= ohello.FOneItem.FSocname_kor %>" disabled></td>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">�귣���(����)</td>
	<td bgcolor="FFFFFF"><input type="text" value="<%= ohello.FOneItem.FSocname %>" disabled></td>
</tr>
<% End If %>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>" >��� BG�̹���</td>
	<td bgcolor="FFFFFF" colspan="3">
		<% If ohello.FOneItem.FBgImageURL <> "" Then %>
		<img src="<%=staticImgUrl%>/brandstreet/hello/<%=ohello.FOneItem.FBgImageURL%>">
			<% If menugubun <> "4" Then %>
		<input type="button" class="button" value="�̹�������" onclick="imgDelProcess();">
			<% End If %>
		<br>
		<% End If %>
		<% if  ohello.FOneItem.FIsSpBrand>0 then  %>
		�귣�� ������ ��� BG�� Ȱ��Ǵ� �̹��� �Դϴ�.<br>(<font color=red>1740 x 668</font> ������� ���ε� ���ּ���)<br>
		<% else %>
		�귣�� ������ ��� BG�� Ȱ��Ǵ� �̹��� �Դϴ�.<br>(1140 x 175 ������� ���ε� ���ּ���)<br>
		<% end if %>
		<input type="file" name="bgImageURL">
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>" >�귣�� �Ұ�<br>(BRAND STORY)</td>
	<td bgcolor="FFFFFF" colspan="3">
		<table border="0" cellpadding="0" cellspacing="0" class="a">
			<tr>
				<td>Ÿ��Ʋ</td>
				<td>&nbsp;&nbsp;<input type="text" class="text" size="100" maxlength=70 name="StoryTitle" value="<%= ohello.FOneItem.FStoryTitle %>">&nbsp;<a onclick="doImgPop('<%=staticImgUrl%>/brandstreet/hello/hello_sample.JPG')" style="cursor:pointer;"><font color="RED">���ú���</font></a></td>
			</tr>
			<tr height="10"><td>&nbsp;</td></tr>
			<tr style="vertical-align:text-top;">
				<td>����</td>
				<td>&nbsp;&nbsp;<textarea name="StoryContent" cols="95" class="textarea" rows="10"><%= ohello.FOneItem.FStoryContent %></textarea></td>
			</tr>
		</table>
	</td>
</tr>
<tr >
	<td align="center" bgcolor="<%= adminColor("tabletop") %>" >�귣�� ö��<br>(PHILOSOPHY)</td>
	<td bgcolor="FFFFFF" colspan="3">
		<table border="0" cellpadding="0" cellspacing="0" class="a">
			<tr>
				<td>Ÿ��Ʋ</td>
				<td>&nbsp;&nbsp;<input type="text" class="text" name="philosophyTitle" size="100" maxlength=70 value="<%= ohello.FOneItem.FPhilosophyTitle %>"></td>
			</tr>
			<tr height="10"><td>&nbsp;</td></tr>
			<tr style="vertical-align:text-top;">
				<td>����</td>
				<td>&nbsp;&nbsp;<textarea cols="95" class="textarea" name="philosophyContent" rows="10"><%= ohello.FOneItem.FPhilosophyContent %></textarea></td>
			</tr>
		</table>
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>" >�������̶�?<br>(Design is?)</td>
	<td bgcolor="FFFFFF" colspan="3">
		�����ο� ���� �귣���� ������ ����ּ���.<br>
		<table border="0" cellpadding="0" cellspacing="0" class="a">
			<tr>
				<td>��������</td>
				<td>&nbsp;<input type="text" class="text" name="designis" size="100" maxlength=70 value="<%= ohello.FOneItem.FDesignis %>"></td>
			</tr>
		</table>
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>" >�귣�� ���ã��</td>
	<td bgcolor="FFFFFF" colspan="3">
		�귣�忡 ������ �ִ� ����ũ�� ����Ʈ�� ������ �ּ���(�ִ� 3�� / �ڻ� ����Ʈ�� ����� ���� ����Ʈ ����)
		<table border="0" cellpadding="0" cellspacing="0" class="a">
			<tr height="60">
				<td>
					1.����Ʈ�� : <input type="text" name="bookmark1SiteName" size="30" maxlength=20 class="text" value="<%= ohello.FOneItem.FBookmark1SiteName %>">
					URL : <input type="text" size="80" maxlength=80 name="bookmark1SiteURL" class="text" value="<%= ohello.FOneItem.FBookmark1SiteURL %>"><br>
					&nbsp;&nbsp;&nbsp;���� : <input type="text" size="90" maxlength=60 name="bookmark1SiteDetail" class="text" value="<%= ohello.FOneItem.FBookmark1SiteDetail %>">
				</td>
			</tr>
		</table>
		<table border="0" cellpadding="0" cellspacing="0" class="a">
			<tr height="60">
				<td>
					2.����Ʈ�� : <input type="text" name="bookmark2SiteName" size="30" maxlength=20 class="text" value="<%= ohello.FOneItem.FBookmark2SiteName %>">
					URL : <input type="text" size="80" maxlength=80 name="bookmark2SiteURL" class="text" value="<%= ohello.FOneItem.FBookmark2SiteURL %>"><br>
					&nbsp;&nbsp;&nbsp;���� : <input type="text" size="90" maxlength=60 name="bookmark2SiteDetail" class="text" value="<%= ohello.FOneItem.FBookmark2SiteDetail %>">
				</td>
			</tr>
		</table>
		<table border="0" cellpadding="0" cellspacing="0" class="a">
			<tr height="60">
				<td>
					3.����Ʈ�� : <input type="text" name="bookmark3SiteName" size="30" maxlength=20 class="text" value="<%= ohello.FOneItem.FBookmark3SiteName %>">
					URL : <input type="text" size="80" maxlength=80 name="bookmark3SiteURL" class="text" value="<%= ohello.FOneItem.FBookmark3SiteURL %>"><br>
					&nbsp;&nbsp;&nbsp;���� : <input type="text" size="90" maxlength=60 name="bookmark3SiteDetail" class="text" value="<%= ohello.FOneItem.FBookmark3SiteDetail %>">
				</td>
			</tr>
		</table>
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>" >�귣�� �±�</td>
	<td bgcolor="FFFFFF" colspan="3">
		<input type="text" class="text" size="95" maxlength=60 name="brandTag" value="<%= ohello.FOneItem.FBrandTag %>">(�޸��� ����)
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>" >���� �귣��<br>(�귣��ID)</td>
	<td bgcolor="FFFFFF" colspan="3">
   		<%
   			If ohello.FOneItem.FSamebrand <> "" Then
   				drawSelectBoxDesignerwithName2 "samebrand", ohello.FOneItem.FSamebrand, ohello.FOneItem.FSamebrand
			Else
				drawSelectBoxDesignerwithName2 "samebrand", ohello.FOneItem.FSamebrand, ""
   			End If
		%>
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>" >�������</td>
	<td bgcolor="FFFFFF" colspan="3">
		<input type="radio" class="radio" name="isusing" value="Y" <%= Chkiif(ohello.FOneItem.FIsusing = "Y", "checked", "") %>>Y
		<input type="radio" class="radio" name="isusing" value="N" <%= Chkiif(ohello.FOneItem.FIsusing = "" OR ohello.FOneItem.FIsusing = "N", "checked", "") %>>N
	</td>
</tr>
<tr>
	<td bgcolor="FFFFFF" align="center" colspan="4" ><input type="button" value="����" class="button" onclick="form_check();"></td>
</tr>
</form>
</table>
<% Set ohello = nothing %>
<iframe name="xLink" id="xLink" frameborder="0" width="0" height="0"></iframe>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ���޸� �Ǹż�������
' Hieditor : ������ ����
'			 2019.07.09 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/classes/etc/outmallConfirmCls.asp"-->
<%
Dim page, i, research
Dim sDt, eDt, oOutmall, currstat, makerid
page		= request("page")
sDt			= request("sDt")
eDt			= request("eDt")
currstat	= request("currstat")
makerid		= request("makerid")
research	= request("research")

If page = "" Then page = 1
If (research = "") Then
	currstat = "1"
End If

SET oOutMall = new cOutmall
	oOutMall.FCurrPage			= page
	oOutMall.FPageSize			= 1000
	oOutMall.FRectmakerid		= makerid
	oOutMall.FRectsDt			= sDt
	oOutMall.FRecteDt			= eDt
	oOutMall.FRectCurrstat		= currstat
	oOutMall.getConfirmList
%>
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<script language='javascript'>
function goPage(pg){
	frm.page.value = pg;
	frm.submit();
}
function popLogView(imallid, imakerid){
    var pwin = window.open("/designer/itemmaster/popHopeLog.asp?mallid="+imallid+"&makerid="+imakerid,"popHopeLog","width=850,height=700,scrollbars=yes,resizable=yes");
	pwin.focus();
}
function popWhyNot(iidx){
    var pwin2 = window.open("/admin/etc/outmall/popWhyNotSell.asp?idx="+iidx,"popWhyNot","width=800,height=300,scrollbars=yes,resizable=yes");
	pwin2.focus();
}
function checkConfirmProcess() {
	var chkSel=0;
	try {
		if(frmSvArr.cksel.length>1) {
			for(var i=0;i<frmSvArr.cksel.length;i++) {
				if(frmSvArr.cksel[i].checked) chkSel++;
			}
		} else {
			if(frmSvArr.cksel.checked) chkSel++;
		}
		if(chkSel<=0) {
			alert("������ �귣�尡 �����ϴ�.");
			return;
		}
	}
	catch(e) {
		alert("�귣�尡 �����ϴ�.");
		return;
	}

	if (confirm('�����Ͻ� ' + chkSel + '�� �귣�带 ���� �Ͻðڽ��ϱ�?')){
		document.getElementById("btnRegSel").disabled=true;
		document.frmSvArr.target = "xLink";
		document.frmSvArr.cmdparam.value = "confirmOK";
		document.frmSvArr.action = "/admin/etc/outmall/confirm_process.asp"
		document.frmSvArr.submit();
    }
}
function popSugiConfirm(){
    var pwin3 = window.open("/admin/etc/outmall/popJaehyu_Not_In_Makerid.asp","popSugiConfirm","width=1200,height=600,scrollbars=yes,resizable=yes");
	pwin3.focus();
}

function itemdisp(makerid){
    var itemdisp = window.open("/admin/itemmaster/itemlist.asp?makerid="+makerid,"itemdisp","width=1600,height=960,scrollbars=yes,resizable=yes");
	itemdisp.focus();
}

/*
function NotInMakerid(imallgubun){
    var pwin4 = window.open("/admin/etc/outmall/popExtUse_Not_In_Makerid.asp?mallgubun="+imallgubun,"popNotInMakerid","width=1200,height=600,scrollbars=yes,resizable=yes");
	pwin4.focus();
}
*/

</script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="page" value="">
<input type="hidden" name="research" value="on">
<input type="hidden" name="menupos" value="<%= menupos %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
		�귣��&nbsp;&nbsp;&nbsp; : <% drawSelectBoxDesignerwithName "makerid",makerid %>&nbsp;
		<br><br>
		��û���� :
		<select name="currstat" class="select">
			<option value="">��ü</option>
			<option value="1" <%= Chkiif(currstat="1","selected","") %> >���δ��</option>
			<option value="3" <%= Chkiif(currstat="3","selected","") %> >���οϷ�</option>
			<option value="2" <%= Chkiif(currstat="2","selected","") %> >�ݷ�</option>
		</select>
		<br><br>
		�����û�� : 
        <input id="sDt" name="sDt" value="<%=sDt%>" class="text" size="10" maxlength="10" />
        <img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="sDt_trigger" border="0" style="cursor:pointer" align="absmiddle" /> ~
        <input id="eDt" name="eDt" value="<%=eDt%>" class="text" size="10" maxlength="10" />
        <img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="eDt_trigger" border="0" style="cursor:pointer" align="absmiddle" />
	</td>
	<td width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="javascript:document.frm.submit();">
	</td>
</tr>
</form>
</table>
<p>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmSvArr" method="post" onSubmit="return false;" action="" style="margin:0px;">
<input type="hidden" name="cmdparam" value="">
<tr height="30" bgcolor="#FFFFFF">
	<td colspan="9">
		<table width="100%" border="0" cellpadding="0" cellspacing="0" class="a">
		<tr>
			<td align="left">
				�˻���� : <b><%= FormatNumber(oOutMall.FTotalCount,0) %></b>
				&nbsp;
				������ : <b> <%= FormatNumber(page,0) %> / <%= FormatNumber(oOutMall.FTotalPage,0) %></b>
			</td>
			<td align="right">
				<!--
				<input type="button" class="button" id="btnSugiRegSel" value="��� ���� �귣��" onclick="NotInMakerid('cjmall');">&nbsp;&nbsp;
				-->
				<input type="button" class="button" id="btnSugiRegSel" value="�������" onclick="popSugiConfirm();">
				<% If currstat = "1" Then %>
				&nbsp;&nbsp;<input type="button" class="button" id="btnRegSel" value="����" onclick="checkConfirmProcess();">
				<% End If %>
			</td>
		</tr>
		</table>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="20"><input type="checkbox" name="chkAll" onClick="fnCheckAll(this.checked,frmSvArr.cksel);"></td>
	<td width="80">�귣��ID</td>
	<td width="100">����</td>
	<td width="140">����Ǹż���</td>
	<td width="140">�����û��</td>
	<td>����</td>
	<td width="80">����</td>
	<td width="70">����Ȯ��</td>
	<td width="80">�ݷ�</td>
</tr>
<% If oOutMall.FResultCount > 0 Then %>
<% For i = 0 To oOutMall.FResultCount - 1 %>
<tr align="center" bgcolor="#FFFFFF">
	<td><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"  value="<%= oOutMall.FItemlist(i).Fidx %>"></td>
	<td width="80">
		<a href="#" onclick="itemdisp('<%= oOutMall.FItemlist(i).FMakerid %>'); return false;">
		<%= oOutMall.FItemlist(i).FMakerid %></a>
	</td>
	<td width="100">
	<%
		Select Case oOutMall.FItemlist(i).FMallgubun
			Case "naverep"			response.write "���̹�"
			Case "daumep"			response.write "����"
			Case "shodocep"			response.write "���"
			Case "all"				response.write "���޻� ��ü"
			Case Else				response.write oOutMall.FItemlist(i).FMallgubun
		End Select
	%>
	</td>
	<td width="140">
	<%
		Select Case oOutMall.FItemlist(i).FHopesellstat
			Case "Y"		response.write "�Ǹ�"
			Case Else		response.write "�Ǹž���"
		End Select
	%>
	</td>
	<td width="140"><%= oOutMall.FItemlist(i).FHoperegdate %></td>
	<td><%= oOutMall.FItemlist(i).FWhyhope %></td>
	<td width="70">
	<%
		Select Case oOutMall.FItemlist(i).FCurrstat
			Case "1"		response.write "���δ��"
			Case "3"		response.write "���οϷ�"
			Case Else		response.write "�ݷ�"
		End Select
	%>
	</td>
	<td width="70"><input type="button" class="button" value="����" onclick="popLogView('<%= oOutMall.FItemList(i).FMallgubun %>', '<%= oOutMall.FItemList(i).FMakerid %>');"></td>
	<td width="70">
	<% If oOutMall.FItemlist(i).FCurrstat = "1" Then %>
		<input type="button" class="button" value="��������" onclick="popWhyNot('<%= oOutMall.FItemList(i).FIdx %>');">
	<% Else %>
		&nbsp;
	<% End If %>
	</td>
</tr>
<% Next %>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="9" align="center">
	<% If oOutMall.HasPreScroll Then %>
		<a href="javascript:goPage('<%= oOutMall.StartScrollPage-1 %>');">[pre]</a>
	<% Else %>
		[pre]
	<% End If %>
	<% For i=0 + oOutMall.StartScrollPage To oOutMall.FScrollCount + oOutMall.StartScrollPage - 1 %>
		<% If i>oOutMall.FTotalpage Then Exit For %>
		<% If CStr(page)=CStr(i) Then %>
		<font color="red">[<%= i %>]</font>
		<% Else %>
		<a href="javascript:goPage('<%= i %>');">[<%= i %>]</a>
		<% End If %>
	<% Next %>
	<% If oOutMall.HasNextScroll Then %>
		<a href="javascript:goPage('<%= i %>');">[next]</a>
	<% Else %>
	[next]
	<% End If %>
	</td>
</tr>
<% Else %>
<tr height="50" bgcolor="FFFFFF">
	<td colspan="9" align="center">
		�����Ͱ� �����ϴ�
	</td>
</tr>
<% End If %>
</table>
<script language="javascript">
	var CAL_Start = new Calendar({
		inputField : "sDt", trigger    : "sDt_trigger",
		onSelect: function() {
			var date = Calendar.intToDate(this.selection.get());
			CAL_End.args.min = date;
			CAL_End.redraw();
			this.hide();
		}, bottomBar: true, dateFormat: "%Y-%m-%d"
	});
	var CAL_End = new Calendar({
		inputField : "eDt", trigger    : "eDt_trigger",
		onSelect: function() {
			var date = Calendar.intToDate(this.selection.get());
			CAL_Start.args.max = date;
			CAL_Start.redraw();
			this.hide();
		}, bottomBar: true, dateFormat: "%Y-%m-%d"
	});
</script>
<iframe name="xLink" id="xLink" frameborder="0" width="100%" height="50"></iframe>
<% SET oOutmall = nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ���޸� �⺻ ����
' Hieditor : ������ ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/classes/etc/outmallConfirmCls.asp"-->
<%
Dim page, i, research, useCheck
Dim oOutmall
page		= request("page")
research	= request("research")

If page = "" Then page = 1

SET oOutMall = new cOutmall
	oOutMall.FCurrPage			= page
	oOutMall.FPageSize			= 1000
	oOutMall.getOutmallSettingList
%>
<script language='javascript'>
function goPage(pg){
	frm.page.value = pg;
	frm.submit();
}
function marginConfirm() {
	var chkSel=0;
	var standardMargin = ""
	try {
		if(frmSvArr.cksel.length>1) {
			for(var i=0;i<frmSvArr.cksel.length;i++) {
				if(frmSvArr.cksel[i].checked) {
					chkSel++;
					standardMargin = standardMargin + frmSvArr.standardMargin[i].value + "*(^!";
				}
			}
		} else {
			if(frmSvArr.cksel.checked){
				 chkSel++;
				 standardMargin = frmSvArr.standardMargin.value;
			}
		}
		if(chkSel<=0) {
			alert("������ ��ǰ�� �����ϴ�.");
			return;
		}
	}
	catch(e) {
		alert("������ ���� �����ϴ�..");
		return;
	}

	if (confirm('�����Ͻ� ' + chkSel + '�� ������ ���� �Ͻðڽ��ϱ�?')){
		document.frmSvArr.target = "xLink";
		document.frmSvArr.cmdparam.value = "marginOK";
		document.frmSvArr.arrstandardMargin.value = standardMargin;
		document.frmSvArr.action = "/admin/etc/outmall/confirm_process.asp"
		document.frmSvArr.submit();
    }
}
function outmallGoodNoMatch(){
	var pCM2 = window.open("/admin/etc/outmall/popGoodNoMatch.asp","popGoodNoMatch","width=1200,height=600,scrollbars=yes,resizable=yes");
	pCM2.focus();
}
</script>
<% If goodNoUpdateUser = "Y" Then %>
<p>
<input type="button" class="button" value="��ǰ�ڵ� ���Ī" onclick="outmallGoodNoMatch();">
<p>
<% End If %>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmSvArr" method="post" onSubmit="return false;" action="" style="margin:0px;">
<input type="hidden" name="cmdparam" value="">
<input type="hidden" name="arrstandardMargin" value="" />
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
				<input type="button" class="button" value="��������" onclick="marginConfirm();">
			</td>
		</tr>
		</table>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="20"><input type="checkbox" name="chkAll" onClick="fnCheckAll(this.checked,frmSvArr.cksel);"></td>
	<td width="100">���޸�ID</td>
	<td width="200">���ظ���</td>
	<td></td>
</tr>
<% If oOutMall.FResultCount > 0 Then %>
<% For i = 0 To oOutMall.FResultCount - 1 %>
<tr align="center" bgcolor="#FFFFFF">
	<td><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"  value="<%= oOutMall.FItemlist(i).FMallid %>"></td>
	<td><%= oOutMall.FItemlist(i).FMallid %></td>
	<td><input type="text" name="standardMargin" value="<%= oOutMall.FItemlist(i).FOutmallstandardMargin %>"></td>
	<td></td>
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
<iframe name="xLink" id="xLink" frameborder="0" width="100%" height="50"></iframe>
<% SET oOutmall = nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
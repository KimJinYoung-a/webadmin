<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/etc/ebay/ebayCls.asp"-->
<%
Dim oEbay, i
Dim cdl, cdm, cds, dspNo
cdl		= requestCheckVar(request("cdl"),3)
cdm		= requestCheckVar(request("cdm"),3)
cds		= requestCheckVar(request("cds"),3)
dspNo	= requestCheckVar(request("dspNo"),20)

If cdl = "" Then
	Call Alert_Close("ī�װ��� �ڵ尡 �����ϴ�.")
	dbget.Close: Response.End
End IF

'// ī�װ��� ���� ����
Set oEbay = new CEbay
	oEbay.FPageSize = 20
	oEbay.FCurrPage = 1
	oEbay.FRectCDL = cdl
	oEbay.FRectCDM = cdm
	oEbay.FRectCDS = cds
	oEbay.getTenEbayCateList

If oEbay.FResultCount <= 0 Then
	Call Alert_Close("�ش� ī�װ��� ������ �����ϴ�.")
	dbget.Close: Response.End
End If
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>

<script language="javascript">
<!--
	// ��Ī �����ϱ�
	function fnSaveForm() {
		var frm = document.frmAct;
		var chkSel=0;
		try {
			if(document.resultFrm.chk.length>1) {
				for(var i=0;i<document.resultFrm.chk.length;i++) {
					if(document.resultFrm.chk[i].checked) chkSel++;
				}
			} else {
				if(document.resultFrm.chk.checked) chkSel++;
			}
			if(chkSel<=0) {
				alert("��Ī�� ī�װ����� ������ �ּ���.");
				return;
			}
		}
		catch(e) {
			alert("��Ī�� ī�װ����� ������ �ּ���.");
			return;
		}

		if(confirm("�����Ͻ� ī�װ����� ��Ī�Ͻðڽ��ϱ�?")) {
			document.resultFrm.action="procEbay.asp";
			document.resultFrm.method="post";
			document.resultFrm.cdl.value = frm.cdl.value;
			document.resultFrm.cdm.value = frm.cdm.value;
			document.resultFrm.cds.value = frm.cds.value;
			document.resultFrm.mode.value ="saveCateArr";
			document.resultFrm.submit();
		}
	}

    function fnDelForm() {
		var frm = document.frmAct;

		if(confirm("���� ��Ī�� ī�װ����� �������� �Ͻðڽ��ϱ�?\n\n�� ��ǰ �Ǵ� ī�װ����� �����Ǵ� ���� �ƴϸ�, ����� ������ �����˴ϴ�.")) {
			frm.mode.value="delCate";
			frm.action="procEbay.asp";
			frm.submit();
		}
	}

	// â�ݱ�
	function fnCancel() {
		if(confirm("�۾��� ����ϰ� â�� �����ðڽ��ϱ�?")) {
			self.close();
		}
	}

	// Ebay ī�װ��� �˻�
	function fnSearchEbayCate(disptpcd) {
	    var srcKwd = document.srcFrm.srcKwd.value;

	    if (srcKwd.length<1) {
	        alert('�˻�� �Է��ϼ���.');
	        document.srcFrm.srcKwd.focus();
	        return;
	    }

	    $.ajax({
    		url: "actFindESMCate.asp?srcKwd="+srcKwd,
    		cache: false,
    		async: false,
    		success: function(message) {
           		$("#cate_result").empty().html(message);
    		},
    		error: function(){
    		    alert(message);
    		}
    	});
	}
//-->
</script>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="#F3F3FF">
<tr valign="top">
	<td><img src="/images/icon_star.gif" align="absbottom">
	<font color="red"><strong>ebay ī�װ��� ��Ī</strong></font></td>
</tr>
</table>
<p>
<!-- ǥ �߰��� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<tr height="5" valign="top">
	<td align="left"><img src="/images/icon_arrow_down.gif" border="0" vspace="5" align="absmiddle"> �ٹ����� ī�װ��� ����</td>
</tr>
</table>
<!-- ǥ �߰��� ��-->
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<tr height="25">
	<td width="80" align="center" bgcolor="#DDDDFF">��з�</td>
	<td bgcolor="#FFFFFF">[<%=cdl%>] <%=oEbay.FItemList(0).FtenCDLName%></td>
</tr>
<tr height="25">
	<td width="80" align="center" bgcolor="#DDDDFF">�ߺз�</td>
	<td bgcolor="#FFFFFF">[<%=cdm%>] <%=oEbay.FItemList(0).FtenCDMName%></td>
</tr>
<tr height="25">
	<td width="80" align="center" bgcolor="#DDDDFF">�Һз�</td>
	<td bgcolor="#FFFFFF">[<%=cds%>] <%=oEbay.FItemList(0).FtenCDSName%></td>
</tr>
</table>

<!-- ǥ �߰��� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<tr height="5" valign="top">
	<td align="left"><img src="/images/icon_arrow_down.gif" border="0" vspace="5" align="absmiddle"> ebay ī�װ��� ��Ī ����</td>
</tr>
</table>
<!-- ǥ �߰��� ��-->
<form name="srcFrm" method="GET" onsubmit="fnSearchEbayCate();return false;" style="margin:0px;">
<input type="hidden" name="srcDiv" value="CNM">
<input type="hidden" name="disptpcd" value="">
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<tr height="25">
    <td id="brTT" width="80" align="center" bgcolor="#DDDDFF" rowspan="2" >�˻�</td>
	<td bgcolor="#FFFFFF">
		ī�װ����� <input type="text" name="srcKwd" class="text">
		<input type="button" value="�˻�" class="button" onClick="fnSearchEbayCate();">
	</td>
</tr>
<tr >
	<td bgcolor="#F2F2F2">
	<div id="cate_result" ></div>
	</td>
</tr>
<tr height="25">
    <td id="brTT" width="80" align="center" bgcolor="#DDDDFF" rowspan="<%= oEbay.FResultCount + 1 %>" >��ϵ�<br>ī�װ���</td>
	<td bgcolor="#FFFFFF" height="1"></td>
</tr>
<% For i = 0 to oEbay.FResultCount - 1 %>
<% If Not IsNULL(oEbay.FItemList(i).FCateCode) Then %>
<tr>
    <td bgcolor="#F2F2F2"><b><span id="selBr"><%= Chkiif(oEbay.FItemList(i).FGubun="A", "<font color='RED'>����</font>", "<font color='GREEN'>������</font>") %>  : <%=oEbay.FItemList(i).FCateName%> [<%= oEbay.FItemList(i).FCateCode%>]</span></b>
    &nbsp;&nbsp;&nbsp;&nbsp;
    </td>
</tr>
<% End If %>
<% Next %>
</table>
</form>
<!-- ǥ �ϴܹ� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<tr valign="top" height="28">
    <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
    <td valign="bottom" align="left"></td>
    <td valign="bottom" align="right">
		<img src="/images/icon_cancel.gif" width="45" height="20" border="0" onclick="fnCancel()" style="cursor:pointer" align="absmiddle"> &nbsp;&nbsp;&nbsp;
		<img src="/images/icon_save.gif" width="45" height="20" border="0" onclick="fnSaveForm()" style="cursor:pointer" align="absmiddle"> &nbsp;&nbsp;&nbsp;
		<% If dspNo <> "" Then %>
		<img src="/images/icon_delete.gif" width="45" height="20" border="0" onclick="fnDelForm();" style="cursor:pointer" align="absmiddle">
		<% End If %>
    </td>
    <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</tr>
</table>
<!-- ǥ �ϴܹ� ��-->
<form name="frmAct" method="POST" style="margin:0px;">
<input type="hidden" name="cdl" value="<%=cdl%>">
<input type="hidden" name="cdm" value="<%=cdm%>">
<input type="hidden" name="cds" value="<%=cds%>">
<input type="hidden" name="depthcode" value="">
<input type="hidden" name="stdcode" value="">
<input type="hidden" name="mode" value="">
<input type="hidden" name="categbn" value="cate">
</form>
<iframe name="xLink" id="xLink" frameborder="0" width="1110" height="110"></iframe>
</p>
<% Set oEbay = Nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
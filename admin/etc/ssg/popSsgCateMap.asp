<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/etc/ssg/ssgItemcls.asp"-->
<%
Dim ossg, i
Dim cdl, cdm, cds, dispNo, isNull4DeprgNm
cdl		= requestCheckVar(request("cdl"),3)
cdm		= requestCheckVar(request("cdm"),3)
cds		= requestCheckVar(request("cds"),3)
dispNo	= requestCheckVar(request("dspNo"),16)

If cdl = "" Then
	Call Alert_Close("ī�װ� �ڵ尡 �����ϴ�.")
	dbget.Close: Response.End
End IF

'// ī�װ� ���� ����
Set ossg = new Cssg
	ossg.FPageSize = 20
	ossg.FCurrPage = 1
	ossg.FRectCDL = cdl
	ossg.FRectCDM = cdm
	ossg.FRectCDS = cds
	ossg.getTenssgCateList

If ossg.FResultCount <= 0 Then
	Call Alert_Close("�ش� ī�װ� ������ �����ϴ�.")
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
				alert("��Ī�� ī�װ��� ������ �ּ���.");
				return;
			}
		}
		catch(e) {
			alert("��Ī�� ī�װ��� ������ �ּ���.");
			return;
		}

		if(confirm("�����Ͻ� ī�װ��� ��Ī�Ͻðڽ��ϱ�?")) {
			document.resultFrm.action="procssg.asp";
			document.resultFrm.method="post";
			document.resultFrm.cdl.value = frm.cdl.value;
			document.resultFrm.cdm.value = frm.cdm.value;
			document.resultFrm.cds.value = frm.cds.value;
			document.resultFrm.mode.value ="saveCateArr";
			document.resultFrm.submit();
		}
	}

    function fnDelForm(stdcode,iDspNo) {
		var frm = document.frmAct;
		if (iDspNo=="") {
		    alert("������ ssg ī�װ��� �������ּ���.");
			return;
		}

		if(confirm("���� ��Ī�� ī�װ��� �������� �Ͻðڽ��ϱ�?\n\n�� ��ǰ �Ǵ� ī�װ��� �����Ǵ� ���� �ƴϸ�, ����� ������ �����˴ϴ�.")) {
			frm.mode.value="delCate";
			frm.stdcode.value=stdcode;
			frm.depthcode.value=iDspNo;
			frm.action="procssg.asp";
			frm.submit();
		}
	}

	// â�ݱ�
	function fnCancel() {
		if(confirm("�۾��� ����ϰ� â�� �����ðڽ��ϱ�?")) {
			self.close();
		}
	}

	// ssg ī�װ� �˻�
	function fnSearchssgCate(disptpcd) {
	    var srcKwd = document.srcFrm.srcKwd.value;
	    if (srcKwd.length<1) {
	        alert('�˻�� �Է��ϼ���.');
	        document.srcFrm.srcKwd.focus();
	        return;    
	    }
	    
	    $.ajax({
    		url: "actFindSsgCate.asp?disptpcd="+disptpcd+"&srcKwd="+srcKwd,
    		cache: false,
    		async: false,
    		success: function(message) {
           		$("#cate_result").empty().html(message);
    		},
    		error: function(){
    		    alert(message);
    		}
    	});
    	
	     
	    
		//var pFCL = window.open("","popssgCate","width=1200,height=700,scrollbars=yes,resizable=yes");
		//pFCL.focus();
		//srcFrm.target="popssgCate";
		//srcFrm.action="popFindssgCate.asp";
		//srcFrm.submit();
	}
//-->
</script>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="#F3F3FF">
<tr valign="top">
	<td><img src="/images/icon_star.gif" align="absbottom">
	<font color="red"><strong>ssg ī�װ� ��Ī</strong></font></td>
</tr>
</table>
<p>
<!-- ǥ �߰��� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<tr height="5" valign="top">
	<td align="left"><img src="/images/icon_arrow_down.gif" border="0" vspace="5" align="absmiddle"> �ٹ����� ī�װ� ����</td>
</tr>
</table>
<!-- ǥ �߰��� ��-->
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<tr height="25">
	<td width="80" align="center" bgcolor="#DDDDFF">��з�</td>
	<td bgcolor="#FFFFFF">[<%=cdl%>] <%=ossg.FItemList(0).FtenCDLName%></td>
</tr>
<tr height="25">
	<td width="80" align="center" bgcolor="#DDDDFF">�ߺз�</td>
	<td bgcolor="#FFFFFF">[<%=cdm%>] <%=ossg.FItemList(0).FtenCDMName%></td>
</tr>
<tr height="25">
	<td width="80" align="center" bgcolor="#DDDDFF">�Һз�</td>
	<td bgcolor="#FFFFFF">[<%=cds%>] <%=ossg.FItemList(0).FtenCDSName%></td>
</tr>
</table>

<!-- ǥ �߰��� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<tr height="5" valign="top">
	<td align="left"><img src="/images/icon_arrow_down.gif" border="0" vspace="5" align="absmiddle"> ssg ī�װ� ��Ī ����</td>
</tr>
</table>
<!-- ǥ �߰��� ��-->
<form name="srcFrm" method="GET" onsubmit="fnSearchssgCate();return false;" style="margin:0px;">
<input type="hidden" name="srcDiv" value="CNM">
<input type="hidden" name="disptpcd" value="">
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<tr height="25">
    <td id="brTT" width="80" align="center" bgcolor="#DDDDFF" rowspan="2" >�˻�</td>
	<td bgcolor="#FFFFFF">
		ī�װ��� <input type="text" name="srcKwd" class="text">
		<input type="button" value="�˻�" class="button" onClick="fnSearchssgCate();">
	</td>
</tr>
<tr >
	<td bgcolor="#F2F2F2">
	<div id="cate_result" ></div>  
	</td>
</tr>
<tr height="25">
    <td id="brTT" width="80" align="center" bgcolor="#DDDDFF" rowspan="<%= ossg.FResultCount + 1 %>" >��ϵ�<br>ī�װ�</td>
	<td bgcolor="#FFFFFF" height="1"></td>
</tr>
<% For i = 0 to ossg.FResultCount - 1 %>
<% If Not IsNULL(ossg.FItemList(i).FDepthCode) Then %>
	<%
		If ossg.FItemList(i).FDepth4Nm = "" Then
			isNull4DeprgNm = ossg.FItemList(i).FDepth3Nm
		Else
			isNull4DeprgNm = ossg.FItemList(i).FDepth4Nm
		End If
	%>
<tr>
    <td bgcolor="#F2F2F2"><b><span id="selBr"><%=ossg.FItemList(i).getSiteNoToSiteName%>  : <%=isNull4DeprgNm%> [<%=ossg.FItemList(i).FStdDepthCode%>] <%=ossg.FItemList(i).FDepthCode%></span></b>
    &nbsp;&nbsp;&nbsp;&nbsp;<img src="/images/icon_delete.gif" width="45" height="20" border="0" onclick="fnDelForm('<%=ossg.FItemList(i).FStdDepthCode%>','<%=ossg.FItemList(i).FDepthCode%>')" style="cursor:pointer" align="absmiddle">
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
		<img src="/images/icon_save.gif" width="45" height="20" border="0" onclick="fnSaveForm()" style="cursor:pointer" align="absmiddle">
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
<input type="hidden" name="mode" value="saveCate">
<input type="hidden" name="categbn" value="cate">
</form>
<iframe name="xLink" id="xLink" frameborder="0" width="1110" height="110"></iframe>
</p>
<% Set ossg = Nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->

<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<%
dim isusing
dim arrItem, arrItemCount, strSql		
	strSql = " SELECT * "
	strSql = strSql & "	FROM db_sitemaster.DBO.tbl_pcmain_top_exhibition_ctrl order by idx asc "	
	
	rsget.CursorLocation = adUseClient
	rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
	
	if Not rsget.Eof Then
	'isQuickDlv = rsget("RESULT")
		arrItem = rsget.GetRows
		arrItemCount = rsget.RecordCount	
	End If
	rsget.close		
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<link href="/js/jqueryui/css/jquery-ui.css" rel="stylesheet">
<script language="javascript">
$(function(){	
    $('#datepicker1').datepicker( {
        changeMonth: true,
        changeYear: true,
        showButtonPanel: true,
        dateFormat: 'yy-mm-dd',     
	});		
    $('#datepicker2').datepicker( {
        changeMonth: true,
        changeYear: true,
        showButtonPanel: true,
        dateFormat: 'yy-mm-dd',     
    });			
});
function previewMain(flatform){
	if(flatform == 'W'){
		if(document.frm.testdate1.value == ""){
			alert("�̸����� ��¥�� �������ּ���.");
			document.frm.testdate1.focus();
			return false;
		}
		var testdate1 = document.frm.testdate1.value
		var url = "<%=vwwwUrl%>?testdate="+testdate1
		window.open(url, "testMain");
	}else{
		if(document.frm.testdate2.value == ""){
			alert("�̸����� ��¥�� �������ּ���.");
			document.frm.testdate2.focus();
			return false;
		}
		var testdate2 = document.frm.testdate2.value
		var winView = window.open("<%=vmobileUrl%>?testdate="+testdate2,"testMain2","width=400, height=600,scrollbars=yes,resizable=yes");
	}

}
function submitCtrlData(){
	var frm = document.frm;
	frm.method = "post"
	frm.action = "event_proc.asp";
	frm.submit();
}
</script>
<%' response.write vwwwUrl%>
<!-- �˻� ���� -->
<form name="frm">
<input type="hidden" name="mode" value="exhibitionOpenCtrl">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>"> 
	<tr align="center" bgcolor="<%= adminColor("topbar") %>" >
		<td rowspan=2 width="70" bgcolor="<%= adminColor("gray") %>">��ȹ�� ����</td>
		<td >
			<div style="float:left;">
			PCWEB : 
				<input type="radio" name="PCisusing" value="1" <%=chkiif(arrItem(1, 0) = "1","checked","")%> checked />����� &nbsp;&nbsp;&nbsp; 
				<input type="radio" name="PCisusing" value="0"  <%=chkiif(arrItem(1, 0) = "0","checked","")%>/>������
			</div> 								
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("topbar") %>" >		
		<td >
			<div style="float:left;">
			MOBILE : 
				<input type="radio" name="mobileisusing" value="1" <%=chkiif(arrItem(1, 1) = "1","checked","")%> checked />����� &nbsp;&nbsp;&nbsp; 
				<input type="radio" name="mobileisusing" value="0"  <%=chkiif(arrItem(1, 1) = "0","checked","")%>/>������
			</div> 								
		</td>
	</tr>	
	<tr align="center" bgcolor="<%= adminColor("topbar") %>" >
		<td  width="70" bgcolor="<%= adminColor("gray") %>">�̸�����<br>pc��</td>
			<td >
			<div style="float:left;">
			��¥:<input type="text" name="testdate1" id="datepicker1" readonly>
				<button type="button" onclick="previewMain('W');">�̸�����</button>
			</div> 
		</td>
	</tr>	
	<tr align="center" bgcolor="<%= adminColor("topbar") %>" >
		<td  width="70" bgcolor="<%= adminColor("gray") %>">�̸�����<br>�����</td>
			<td >
			<div style="float:left;">
			��¥:<input type="text" name="testdate2" id="datepicker2" readonly>
				<button type="button" onclick="previewMain('M');">�̸�����</button>
			</div> 
		</td>
	</tr>		
	<tr bgcolor="#FFFFFF" align="center">
		<td colspan="2">
			<input type="button" value=" �� �� " onClick="submitCtrlData();"/>
			<input type="button" value=" �� �� " onClick="window.close();"/>			
		</td>
	</tr>	
</table>
</form>
<script type="text/javascript" src="/js/jquery-ui-1.10.3.custom.min.js"></script>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<%
	dim realdate : realdate = request("realdate")
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
</script>
<%' response.write vwwwUrl%>
<!-- �˻� ���� -->
<span style="color:red">* �ٹ����ٿ��� �α����� �ϼž� �̸����� ����� ����Ͻ� �� �ֽ��ϴ�. </span>
<form name="frm">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>"> 
	<tr align="center" bgcolor="<%= adminColor("topbar") %>" >
		<td  width="70" bgcolor="<%= adminColor("gray") %>">�̸�����<br>pc��</td>
			<td >
			<div style="float:left;">
			��¥:<input type="text" name="testdate1" id="datepicker1" readonly value="<%=chkiif(realdate <> "" ,realdate , "")%>">
				<button type="button" onclick="previewMain('W');">�̸�����</button>
			</div> 
		</td>
	</tr>	
	<!--<tr align="center" bgcolor="<%= adminColor("topbar") %>" >
		<td  width="70" bgcolor="<%= adminColor("gray") %>">�̸�����<br>�����</td>
			<td >
			<div style="float:left;">
			��¥:<input type="text" name="testdate2" id="datepicker2" readonly>
				<button type="button" onclick="previewMain('M');">�̸�����</button>
			</div> 
		</td>
	</tr>-->
</table>
</form>
<script type="text/javascript" src="/js/jquery-ui-1.10.3.custom.min.js"></script>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
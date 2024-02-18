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
			alert("미리보기 날짜를 설정해주세요.");
			document.frm.testdate1.focus();
			return false;
		}
		var testdate1 = document.frm.testdate1.value
		var url = "<%=vwwwUrl%>?testdate="+testdate1
		window.open(url, "testMain");
	}else{
		if(document.frm.testdate2.value == ""){
			alert("미리보기 날짜를 설정해주세요.");
			document.frm.testdate2.focus();
			return false;
		}
		var testdate2 = document.frm.testdate2.value
		var winView = window.open("<%=vmobileUrl%>?testdate="+testdate2,"testMain2","width=400, height=600,scrollbars=yes,resizable=yes");
	}

}
</script>
<%' response.write vwwwUrl%>
<!-- 검색 시작 -->
<span style="color:red">* 텐바이텐에서 로그인을 하셔야 미리보기 기능을 사용하실 수 있습니다. </span>
<form name="frm">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>"> 
	<tr align="center" bgcolor="<%= adminColor("topbar") %>" >
		<td  width="70" bgcolor="<%= adminColor("gray") %>">미리보기<br>pc웹</td>
			<td >
			<div style="float:left;">
			날짜:<input type="text" name="testdate1" id="datepicker1" readonly value="<%=chkiif(realdate <> "" ,realdate , "")%>">
				<button type="button" onclick="previewMain('W');">미리보기</button>
			</div> 
		</td>
	</tr>	
	<!--<tr align="center" bgcolor="<%= adminColor("topbar") %>" >
		<td  width="70" bgcolor="<%= adminColor("gray") %>">미리보기<br>모바일</td>
			<td >
			<div style="float:left;">
			날짜:<input type="text" name="testdate2" id="datepicker2" readonly>
				<button type="button" onclick="previewMain('M');">미리보기</button>
			</div> 
		</td>
	</tr>-->
</table>
</form>
<script type="text/javascript" src="/js/jquery-ui-1.10.3.custom.min.js"></script>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
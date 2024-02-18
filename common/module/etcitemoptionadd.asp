<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/common/incSessionBctId.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<%
'###########################################################
' Description : 상품 전용옵션 등록
' History : 2013.12.16 정윤정 옵션갯수 수정   
'###########################################################
%>
<%
dim i,iRowMax
iRowMax = 19 '옵션 최대갯수
%>

<script language="javascript">
<!--
function AddOption()
{
	var frm = document.itemopt;
    var addedCnt = 0;
    
	if(!frm.optTypeNm.value){
		alert("추가할 옵션 구분 명을 입력해주십시오.");
		frm.optTypeNm.focus();
		return false;
	}

    for (var i=0;i<frm.optNm.length;i++){
        if (frm.optNm[i].value.length>0){
            opener.InsertOptionWithGubun(frm.optTypeNm.value, frm.optNm[i].value, "0000");
            addedCnt++;
        }
    }

    if (addedCnt>0){
	    self.close();
	}else{
	    alert('추가할 옵션을 입력해 주세요.');
	}
}
//-->
</script>
<body onload="window.resizeTo(550,890);document.itemopt.optTypeNm.focus();">
<table width="500" border="0" cellspacing="1" cellpadding="2" align="center" class="a"  bgcolor="#3d3d3d">
<form name="itemopt" >
    <tr height="30" bgcolor="#DDDDFF">
		<td width="120" align="center">옵션 구분 명</td>
		<td bgcolor="#FFFFFF" align="left"><input type="text" name="optTypeNm" size="20" maxlength="20"> 색상</td>
	</tr>
	<% for i=0 to iRowMax %>
	<tr height="30" bgcolor="#DDDDFF">
		<td width="120" align="center">옵션 명 <%= i+1 %></td>
		<td bgcolor="#FFFFFF" align="left"><input type="text" name="optNm" size="32" maxlength="20"> <%= chkIIF(i=0,"빨강","") %><%= chkIIF(i=1,"파랑","") %><%= chkIIF(i=2,"노랑","") %></td>
	</tr>
	<% next %>
	<tr bgcolor="#FFFFFF">
		<td colspan="2" align="center">
			<input type="button" value="옵션추가" class="button" onClick="AddOption();">
			<input type="button" value=" 닫 기 "  class="button" onclick="self.close()">
		</td>
	</tr>
</form> 
</table>
<div style="padding:5px;text-align:right;font-size:8pt">Ver1.0  lastupdate: 2013.12.16 </div>
</body>
<!-- #include virtual="/admin/lib/poptail.asp"-->
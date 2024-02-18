<%@ language=vbscript %>
<% option explicit %>

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/RedRibbon/redRibbonManagerCls.asp"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<script language="JavaScript" src="/js/xl.js"></script>
<script language="JavaScript" src="/js/common.js"></script>
<script language="JavaScript" src="/js/report.js"></script>
<link rel="stylesheet" href="/css/scm.css" type="text/css">
</head>
<body>

<%

dim cdL,cdM,cdS

cdL= request("cdL")
cdM= request("cdM")
cdS= request("cdS")

dim objView

set objView = new giftManagerView
objView.getMenuView cdL,cdM,cdS


%>
<script language="javascript">

document.domain = "10x10.co.kr";

function subchk(){
	
	var min = document.getElementsByName("minvalue");
	var max = document.getElementsByName("maxvalue");
	
	for (var i=0;i<min.length;1++){
		if (isNaN(min[i].value)){
			alert("숫자만 입력하셔야합니다");
			return false;
		}
		
	}
	
	
	for (var i=0;i<max.length;1++){
		if (isNaN(max[i].value)){
			alert("숫자만 입력하셔야합니다");
			return false;
		}
	}
	return;
}

// 검색조건 추가
function addSearhchCash(min,max){
	
	var tbl = document.getElementById("cashtbl");
	if (tbl.rows.length < 4){
		var oRow = tbl.insertRow();
		var oCell = oRow.insertCell();
		oCell.align="center";
		oCell.innerHTML = "<input class='input_a' type='text' size='12' name='minvalue' value='" + min + "'>이상 ~ <input class='input_a' type='text' size='12' name='maxvalue' value='" + max +"'>미만 <input type='button' class='button' value='x' onclick='delSearchCash(parentElement.parentElement.rowIndex);'>";
		
	} else {
	
		alert("검색가격은 4개까지만 가능합니다");
	}
}
//검색 조건 삭제
function delSearchCash(num){
	var tbl = document.getElementById("cashtbl");
		
		
	if (tbl.rows.length <= 1){
		alert("하나이상은 입력하셔야 합니다");
		return;
	}else{
		tbl.deleteRow(num);
	}

}
//기본 검색가 생성
function basicSearchCash(){
	var tbl = document.getElementById("cashtbl");
	
	for(var i= tbl.rows.length-1 ; i>=0 ; i--){
		tbl.deleteRow(i);
	}
	
	addSearhchCash('','30000');
	addSearhchCash('30000','60000');
	addSearhchCash('60000','90000');
	addSearhchCash('90000','');
}


</script>

<table width="400" border="0" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="UpdateFRM" action="Menu_Process.asp" target="frame" onsubmit="return subchk();">
	<input type="hidden" name="mode" value="cashedit" />
	<input type="hidden" name="LCode" value="<%= objView.LCode %>" />
	<input type="hidden" name="MCode" value="<%= objView.MCode %>" />
	<input type="hidden" name="SCode" value="<%= objView.SCode %>" />
	<tr>
		<td width="100" bgcolor="#FFFFFF"></td>
		<td width="280" bgcolor="#FFFFFF" align="center">가격설정</td>
	</tr>
<% IF objView.LCode <>"" then %>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>" align="center">대 카테고리</td>
		<td bgcolor="#FFFFFF"> [<font color="red"><%= objView.LCode %></font>] <%= objView.LCodeNm %>
	</tr>
<% END IF %>
<% IF objView.MCode <>"" then %>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>" align="center">중 카테고리</td>
		<td bgcolor="#FFFFFF"> [<font color="red"><%= objView.MCode %></font>] <%= objView.MCodeNm %>
	</tr>
<% END IF %>
<% IF objView.SCode <>"" then %>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>" align="center">소 카테고리</td>
		<td bgcolor="#FFFFFF"> [<font color="red"><%= objView.SCode %></font>] <%= objView.SCodeNm %>
	</tr>
<% END IF %>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>" align="center">가격</td>
		<td bgcolor="#FFFFFF" align="center">
			<table width="100%" border="0" cellpadding="0" cellspacing="0" class="a" id="cashtbl"> 
				<% 
				dim strSQL
				strSQL =" SELECT MinCash,MaxCash " &_
						" FROM [db_redribbon].[dbo].[tbl_redR_CashMenu] " &_
						" WHERE LCode='" & cdL &"' "
						
						IF cdM="" THEN
							strSQL = strSQL & " and MCode is Null "
						ELSE
							strSQL = strSQL & " and MCode ='" & cdM & "'"
						END IF
						
						IF cdS="" THEN
							strSQL = strSQL & " and SCode is Null "
						ELSE
							strSQL = strSQL & " and SCode ='" & cdS & "'"
						END IF
				
				rsget.open strSQL,dbget,1
				if not rsget.eof then
				do until rsget.eof
				%>
				<tr>
					<td align="center">
				
						<input type="text" size="12" class="input_a" name="minvalue" value="<%= rsget("minCash") %>">이상
						 ~ 
						<input type="text" size="12" class="input_a" name="maxvalue" value="<%= rsget("maxCash") %>">미만
						<input type="button" class="button" value="x" onclick="delSearchCash(parentElement.parentElement.rowIndex);">
						
					</td>
				</tr>
				<%
				rsget.movenext
				loop
				else
				%>
				<tr>
					<td align="center">
				
						<input type="text" size="12" class="input_a" name="minvalue" value="">이상
						 ~ 
						<input type="text" size="12" class="input_a" name="maxvalue" value="">미만
						<input type="button" class="button" value="x" onclick="delSearchCash(parentElement.parentElement.rowIndex);">
						
					</td>
				</tr>
				<%
				end if
				rsget.close
				%>
			</table>	
		</td>
	</tr>
	<tr>
		<td colspan="2" bgcolor="#FFFFFF" align="center">
			<input type="button" class="button" value="추가" onclick="addSearhchCash('','');">
			<input type="button" class="button" value="기본검색가 생성" onclick="basicSearchCash();">
			<input type="submit" class="button" value="저장">
		</td>
	</tr>
	<tr>
		<td colspan="2" bgcolor="#FFFFFF" align="center"></td>
	</tr>
	
	</form>
</table>

<iframe name="frame" src="" frameborder="0" width="0" height="0"></iframe>
<% set objView = nothing %>
</body>
</html>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
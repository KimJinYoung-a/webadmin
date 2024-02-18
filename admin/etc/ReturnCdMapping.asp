<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<%
Dim popmid, makerid, ckRadio, mode, gCode
popmid	= request("popmid")
makerid = request("makerid")
ckRadio = request("ckRadio")
mode 	= request("mode")

Dim ReturnName, ReturnCode, mappidNullYN, groupCode, sWhere
ReturnName 		= request("ReturnName")
ReturnCode 		= request("ReturnCode")
mappidNullYN	= request("mappidNullYN")
groupCode		= request("groupCode")

If ReturnName <> "" Then
	sWhere = sWhere & " AND r.ReturnName = '"&ReturnName&"' "
End If

If ReturnCode <> "" Then
	sWhere = sWhere & " AND r.ReturnCode = '"&ReturnCode&"' "
End If

If mappidNullYN = "Y" Then
	sWhere = sWhere & " AND r.mapPid IS NOT NULL "
ElseIf mappidNullYN = "N" Then
	sWhere = sWhere & " AND r.mapPid IS NULL "	
End If

If groupCode = "Y" Then
	Dim sql
	sql = ""
	sql = " SELECT groupid FROM db_partner.dbo.tbl_partner WHERE id = '"&popmid&"' " & VBCRLF
	rsget.Open sql,dbget,1
		IF not rsget.EOF THEN
			gCode = rsget("groupid")
		END IF	
	rsget.close

	sWhere = sWhere & " AND p.groupid = '"&gCode&"' "
End If

Dim strSql, strSql2, arrList, i
strSql = ""
strSql = strSql & " SELECT r.mallgubun, r.ReturnCode, r.ReturnName, r.ReturnAddress, r.mapPid, p.id, p.groupid " & VBCRLF
strSql = strSql & " FROM db_temp.dbo.tbl_jaehyumall_returnInfo AS R " & VBCRLF
strSql = strSql & " LEFT JOIN db_partner.dbo.tbl_partner AS p ON r.mappid = p.id " & VBCRLF
strSql = strSql & " WHERE 1=1 "&sWhere&" " & VBCRLF
strSql = strSql & " ORDER BY r.mappid ASC "
'strSql = strSql & " FROM db_temp.dbo.tbl_jaehyumall_returnInfo where 1=1 "&sWhere&"  " & VBCRLF
rsget.Open strSql,dbget,1
	IF not rsget.EOF THEN
		arrList = rsget.getRows() 
	END IF	
rsget.close

If mode = "Y" Then
	strSql2 = ""
	strSql2 = strSql2 & " UPDATE db_temp.dbo.tbl_jaehyumall_returnInfo SET " & VBCRLF
	strSql2 = strSql2 & " mapPid = '"&makerid&"' " & VBCRLF
	strSql2 = strSql2 & " WHERE ReturnCode = '"&ckRadio&"' AND ReturnName = '"&makerid&"' AND mallgubun = 'lotteCom' " & VBCRLF
	dbget.Execute strSql2

	strSql2 = ""
	strSql2 = strSql2 & " INSERT INTO db_item.dbo.tbl_OutMall_BrandReturnCode " & VBCRLF
	strSql2 = strSql2 & " (mallid, makerid, returnCode) " & VBCRLF
	strSql2 = strSql2 & " VALUES " & VBCRLF
	strSql2 = strSql2 & " ('lotteCom', '"&makerid&"', '"&ckRadio&"') " & VBCRLF
	dbget.Execute strSql2
End If
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<script language="JavaScript" src="/js/common.js"></script>
<link rel="stylesheet" href="/css/scm.css" type="text/css">
<script>
function form_check(objName){
	var frm;
	var isChecked = false;
	frm = document.frmSubmit;
	var radioObj = frm.all(objName);

	if(!radioObj.length){
		if(frm.ckRadio.checked == false){
			alert('라디오버튼을 선택하세요');
			return;			
		}else{
			isChecked = true;
		}
	}else{
		for(i=0; i<radioObj.length; i++){
			if(radioObj[i].checked){
				isChecked = true;
				break;
			}
		}
	}
	if(isChecked == false){
		alert('라디오버튼을 선택하세요');
		return;
	}

	if(confirm("정말로 맵핑하시겠습니까?") == true){
		frm.submit();
	}else{
		return false; 
	}
}
</script>
</head>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmsearch" method="get" action="<%=CurrURL()%>" style="margin:0px;">
<input type="hidden" name="makerid" value="<%=popmid%>">
<tr bgcolor="#FFFFFF">
	<td height="50">
		<table width="100%" class="a">
		<tr>
			<td>
				★브랜드명 : <input type="text" class="text" name="ReturnName" value="<%=ReturnName%>" size="20">
				&nbsp;&nbsp;
				★반품코드 : <input type="text" class="text" name="ReturnCode" value="<%=ReturnCode%>" size="20">
				<br>
				★비매칭 브랜드 : <input type="text" name="popmid" class="text" value="<%=popmid%>" readonly size="20" style="background-color:#D9E5FF">
				&nbsp;&nbsp;
				★기매핑여부 : <select class="select" name="mappidNullYN">
				    <option value="" <%=chkiif(mappidNullYN = "","selected","")%> >전체</option>
					<option value="Y" <%=chkiif(mappidNullYN = "Y","selected","")%> >Y</option>
					<option value="N" <%=chkiif(mappidNullYN = "N","selected","")%> >N</option>
				</select>
				&nbsp;&nbsp;
				<input type="checkbox" class="checkbox" name="groupCode" value="Y" <%=chkiif(groupCode = "Y","checked","")%>>동일그룹코드(<%=gCode%>)
			</td>
			<td><input type="submit" value="검 색" style="width:50px;height:50px;"></td>
		</tr>
		</table>
	</td>
</tr>
</form>
</table>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmSubmit" method="post" action="ReturnCdMapping.asp">
<input type="hidden" name="makerid" value="<%=popmid%>">
<input type="hidden" name="mode" value="Y">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="30">
	<td>&nbsp;</td>
	<td>제휴몰</td>
	<td>반품코드</td>
	<td>브랜드명</td>
	<td>MapPid</td>
	<td>반품주소</td>
	<td>기매핑여부</td>
</tr>
<%
	IF isArray(arrList) THEN
		For i =0 To UBound(arrList,2)
%>
<tr align="center" bgcolor="#FFFFFF" height="30">
	<td><input type="radio" name="ckRadio" value="<%=arrList(1,i)%>"></td>
	<td><%=arrList(0,i)%></td>
	<td><%=arrList(1,i)%></td>
	<td><%=arrList(2,i)%></td>
	<td><%=arrList(4,i)%></td>
	<td><%=arrList(3,i)%></td>
	<td><%=chkiif(ISNULL(arrList(4,i)),"N","Y")%></td>
</tr>
<%
		Next
	End If
%>
<tr bgcolor="#FFFFFF" align="center">
	<td colspan="7"><input type="button" class="button" value="확인" onclick="javascript:form_check('ckRadio');"></td>
</tr>
</form>
</table>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
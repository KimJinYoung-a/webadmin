<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/etc/JaehyureturnCodecls.asp"-->
<%
Dim mallgubun, i, makerid, lotteSellyn, regCntYN, addrChk, notMakerId, lo
mallgubun 	= request("mallgubun")
makerid 	= request("makerid")
lotteSellyn = request("lotteSellyn")
regCntYN	= request("regCntYN")
addrChk		= request("addrChk")
notMakerId	= request("notMakerId")
lo			= request("lo")

Dim currPage, TotalCount, PageSize
	PageSize = 15
	currPage		= NullFillWith(Request("cp"),1)

If currPage = ""	 Then currPage = 1
If lo = ""			 Then lo = 1

Dim ReturnCode
SET ReturnCode = new RtCodeList
	ReturnCode.FPageSize	= PageSize
	ReturnCode.FCurrPage	= currPage
	ReturnCode.FMakerid		= makerid
	ReturnCode.FAddrChk		= addrChk
	ReturnCode.NotRtCodeList
	TotalCount = ReturnCode.FTotalCount
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<script language="JavaScript" src="/js/xl.js"></script>
<script language="JavaScript" src="/js/common.js"></script>
<script language="JavaScript" src="/js/report.js"></script>
<link rel="stylesheet" href="/css/scm.css" type="text/css">
</head>
<script language="javascript">
function jsGoPage(iP){
	document.frmpage.cp.value = iP;
	document.frmpage.submit();
}
function GoRt(lo){
	if(lo == "1"){
		location.replace('/admin/etc/JaehyuMall_ReturnCode_Mappid.asp?mallgubun=lotteCom&lo=1');
	}else if(lo == "2"){
		location.replace('/admin/etc/JaehyuMall_ReturnCode_Not_Mappid.asp?mallgubun=lotteCom&lo=2');
	}
}
</script>
<body onload="javascript:window.resizeTo(1500, 800);">
<center>Mall 구분 : <b><%=mallgubun%></b></center>
<form name="frmsearch" method="get" action="<%=CurrURL()%>" style="margin:0px;">
<input type="hidden" name="mallgubun" value="<%=mallgubun%>">
<input type="hidden" name="lo" value="<%=lo%>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="#FFFFFF">
	<td>
		<select class="select" onchange="GoRt(this.value);">
			<option value="1" <%=chkiif(lo = "1","selected","")%>>맵핑완료브랜드</option>
			<option value="2" <%=chkiif(lo = "2","selected","")%>>맵핑코드별</option>
		</select>
	</td>
	<td align="right"><a href="/admin/etc/apiReturnCdReload.asp" target="ifrm"><img src="/images/button_reload.gif" style="cursor:pointer;" border="0"></a></td>
</tr>
</table>
<iframe name="ifrm" id="ifrm" width="0" height="0" frameborder=0></iframe>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="#FFFFFF">
	<td height="50">
		<table width="100%" class="a">
		<tr>
			<td>
				★브랜드ID : <input type="text" class="text" name="makerid" value="<%=makerid%>" size="20"> <input type="button" class="button" value="ID검색" onclick="jsSearchBrandID(this.form.name,'makerid');" >
				&nbsp;&nbsp;
				★관리 : <select class="select" name="addrChk">
					<option value="" selected>전체</option>
					<option value="O" <%=chkiif(addrChk = "O","selected","")%> >완료</option>
					<option value="X" <%=chkiif(addrChk = "X","selected","")%> >맵핑필요</option>
				</select>
			</td>
			<td><input type="submit" value="검 색" style="width:50px;height:50px;"></td>
		</tr>
		</table>
	</td>
</tr>
</table>
</form>
검색 수 : <strong><%=TotalCount%></strong>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="30">
	<td>브랜드ID</td>
	<td>상품수</td>
	<td>반품코드</td>
	<td>파트너전화번호</td>
	<td>파트너반품주소</td>
	<td>제휴몰반품주소</td>
	<td>관리</td>
</tr>
<% For i = 0 to ReturnCode.FResultCount - 1 %>
<tr align="center" bgcolor="#FFFFFF" height="30">
	<td><%=ReturnCode.FItemList(i).FMakerId%></td>
	<td><%=ReturnCode.FItemList(i).FCNT%></td>
	<td><%=ReturnCode.FItemList(i).FReturnCode%></td>
	<td><%=ReturnCode.FItemList(i).FDeliver_phone%></td>
	<td><%=ReturnCode.FItemList(i).FMapAddress%></td>
	<td><%=ReturnCode.FItemList(i).FReturnAddress%></td>
	<td><% If IsNull(ReturnCode.FItemList(i).FReturnCode) Then %><input type="button" class="button" value="맵핑" onclick="window.open('/admin/etc/ReturnCdMapping.asp?popmid=<%=ReturnCode.FItemList(i).FMakerId%>');"><% Else response.write "<font color='BLUE'>완료</font>" End If %></td>
</tr>
<% Next %>
</table>
<form name="frmpage" method="get" action="<%=CurrURL()%>" style="margin:0px;">
<input type="hidden" name="cp" value="<%=currPage%>">
<input type="hidden" name="mallgubun" value="<%=mallgubun%>">
<input type="hidden" name="makerid" value="<%=makerid%>">
<input type="hidden" name="addrChk" value="<%=addrChk%>">
<input type="hidden" name="lo" value="<%=lo%>">

<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("tablebg") %>">
<%
Dim iStartPage, iEndPage, ix, iTotalPage
iStartPage = (Int((currPage-1)/10)*10) + 1
iTotalPage 	=  int((TotalCount-1)/PageSize) +1

If (currPage mod PageSize) = 0 Then
	iEndPage = currPage
Else
	iEndPage = iStartPage + (10-1)
End If
%>
<tr bgcolor="FFFFFF">
	<td height="30" align="center">
		<% If (iStartPage-1 )> 0 Then %><a href="javascript:jsGoPage(<%= iStartPage-1 %>)" onfocus="this.blur();">[pre]</a>
		<% Else %>[pre]<% End If %>
        <%
			For ix = iStartPage to iEndPage
				If (ix > iTotalPage) then Exit For
				If Cint(ix) = Cint(currPage) Then
		%>
			<a href="javascript:jsGoPage(<%= ix %>)" class="menu_link3" onfocus="this.blur();"><font color="red">[<%=ix%>]</font></a>
		<%		Else %>
			<a href="javascript:jsGoPage(<%= ix %>)" class="menu_link3" onfocus="this.blur();">[<%=ix%>]</a>
		<%
				End If
			Next
		%>
    	<% If Cint(iTotalPage) > Cint(iEndPage) Then %><a href="javascript:jsGoPage(<%= ix %>)" onfocus="this.blur();">[next]</a>
		<% Else %>[next]<% End If %>
	</td>
</tr>
</table>
</form>
</body>
</html>
<% SET ReturnCode = nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/etc/JaehyureturnCodecls.asp"-->
<%
Dim mallgubun, i, makerid, lotteSellyn, regCntYN, addrChk, notMakerId, isusing, lo
mallgubun 		= request("mallgubun")
makerid 		= request("makerid")
lotteSellyn 	= request("lotteSellyn")
regCntYN		= request("regCntYN")
addrChk			= request("addrChk")
notMakerId		= request("notMakerId")
isusing			= request("isusing")
lo				= request("lo")

Dim currPage, TotalCount, PageSize
	PageSize = 15
	currPage		= NullFillWith(Request("cp"),1)

If currPage = ""	Then currPage = 1
If lo = ""			Then lo = 1

Dim ReturnCode
SET ReturnCode = new RtCodeList
	ReturnCode.FPageSize		= PageSize
	ReturnCode.FCurrPage		= currPage
	ReturnCode.FMakerid			= makerid
	ReturnCode.FLotteSellyn		= lotteSellyn
	ReturnCode.FAddrChk			= addrChk
	ReturnCode.FRegCntYN		= regCntYN
	ReturnCode.FNotMakerId		= notMakerId
	ReturnCode.FIsusing			= isusing
	ReturnCode.RtCodeList
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
<center>Mall ���� : <b><%=mallgubun%></b></center>
<form name="frmsearch" method="get" action="<%=CurrURL()%>" style="margin:0px;">
<input type="hidden" name="mallgubun" value="<%=mallgubun%>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="#FFFFFF">
	<td>
		<select class="select" onchange="GoRt(this.value);">
			<option value="1" <%=chkiif(lo = "1","selected","")%>>���οϷ�귣��</option>
			<option value="2" <%=chkiif(lo = "2","selected","")%>>�����ڵ庰</option>
		</select>
	</td>
</tr>
</table>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="#FFFFFF">
	<td height="50">
		<table width="100%" class="a">
		<tr>
			<td>
				�ں귣��ID : <input type="text" class="text" name="makerid" value="<%=makerid%>" size="20"> <input type="button" class="button" value="ID�˻�" onclick="jsSearchBrandID(this.form.name,'makerid');" >
				&nbsp;&nbsp;
				�����޸� ��ǰ �Ǹſ��� (��ü��ۻ�ǰ): <select class="select" name="lotteSellyn">
				    <option value="" <%=chkiif(lotteSellyn = "","selected","")%> >��ü</option>
					<option value="Y" <%=chkiif(lotteSellyn = "Y","selected","")%> >Y</option>
					<option value="N" <%=chkiif(lotteSellyn = "N","selected","")%> >N</option>
				</select>
				&nbsp;&nbsp;
				���ּ�üũ : <select class="select" name="addrChk">
					<option value="" selected>��ü</option>
					<option value="O" <%=chkiif(addrChk = "O","selected","")%> >O</option>
					<option value="X" <%=chkiif(addrChk = "X","selected","")%> >X</option>
				</select>
				&nbsp;&nbsp;
				�ڻ�ǰ�� 0 : <input type="checkbox" name="regCntYN" value="Y" <%=chkiif(regCntYN = "Y","checked","")%>>
				&nbsp;&nbsp;
				�ڵ�����ܺ귣�� : <input type="checkbox" name="notMakerId" value="Y" <%=chkiif(notMakerId = "Y","checked","")%>>
				<br>
				�ڷԵ����λ�뿩�� : <select class="select" name="isusing">
				    <option value="" <%=chkiif(isusing = "","selected","")%> >��ü</option>
					<option value="Y" <%=chkiif(isusing = "Y","selected","")%> >Y</option>
					<option value="N" <%=chkiif(isusing = "N","selected","")%> >N</option>
				</select>
			</td>
			<td><input type="submit" value="�� ��" style="width:50px;height:50px;"></td>
		</tr>
		</table>
	</td>
</tr>
</table>
</form>
<br>
�˻� �� : <strong><%=TotalCount%></strong>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="30">
	<td>�귣���</td>
	<td>������ܺ귣��</td>
	<td>���޸���ǰ�ڵ�</td>
	<td>�ּ�üũ</td>
	<td>�귣��ID</td>
	<td>���ι�ǰ�ڵ�</td>
	<td>��Ʈ�ʹ�ǰ�ּ�</td>
	<td>���޸���ǰ�ּ�</td>
	<td>�Ե����λ�뿩��</td>
	<td>��ǰ��</td>
</tr>
<% For i = 0 to ReturnCode.FResultCount - 1 %>
<tr align="center" bgcolor="#FFFFFF" height="30">
	<td><%=ReturnCode.FItemList(i).FReturnName%></td>
	<td><font color="<%=chkIIF(IsNull(ReturnCode.FItemList(i).FNotMakerId),"BLUE","RED")%>"><%=chkiif(IsNull(ReturnCode.FItemList(i).FNotMakerId),"�ش����","�ش�")%></font></td>
	<td><%=ReturnCode.FItemList(i).FTempReturnCode%></td>
	<td><font color="<%=chkIIF(ReturnCode.FItemList(i).FAddrChk = 1,"BLUE","RED")%>"><%=chkIIF(ReturnCode.FItemList(i).FAddrChk = 1,"O","X")%></font></td>
	<td><%=ReturnCode.FItemList(i).FMakerId%></td>
	<td><%=ReturnCode.FItemList(i).FReturnCode%></td>
	<td><%=ReturnCode.FItemList(i).FMapAddress%></td>
	<td><%=ReturnCode.FItemList(i).FReturnAddress%></td>
	<td><font color="<%=chkIIF(ReturnCode.FItemList(i).FIsusing = "Y","BLUE","RED")%>"><%=ReturnCode.FItemList(i).FIsusing%></font></td>
	<td><%=ReturnCode.FItemList(i).FRegCnt%></td>
</tr>
<% Next %>
</table>
<form name="frmpage" method="get" action="<%=CurrURL()%>" style="margin:0px;">
<input type="hidden" name="cp" value="<%=currPage%>">
<input type="hidden" name="mallgubun" value="<%=mallgubun%>">
<input type="hidden" name="makerid" value="<%=makerid%>">
<input type="hidden" name="lotteSellyn" value="<%=lotteSellyn%>">
<input type="hidden" name="addrChk" value="<%=addrChk%>">
<input type="hidden" name="regCntYN" value="<%=regCntYN%>">
<input type="hidden" name="notMakerId" value="<%=notMakerId%>">
<input type="hidden" name="isusing" value="<%=isusing%>">
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
<%
SET ReturnCode = nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
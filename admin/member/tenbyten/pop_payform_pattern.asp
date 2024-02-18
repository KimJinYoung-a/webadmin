<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  사원별 계약정보 패턴
' History : 2011.01.07  정윤정  생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenMemberCls.asp" -->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenPayCls.asp" -->
<%
Dim clsPayForm
Dim arrList, intLoop
Dim part_sn, sRectPatternName
Dim iPageSize,iCurrpage
Dim iTotCnt, iTotalPage
Dim sEmpno,iDefaultPaySeq
dim ino
iDefaultPaySeq =requestCheckvar(request("iDPS"),10)
sEmpno =   requestCheckvar(request("sEN"),14)
iCurrpage =   requestCheckvar(request("iCP"),10)
part_sn=   requestCheckvar(request("part_sn"),10)
sRectPatternName=   requestCheckvar(request("sRPN"),60)
ino =requestCheckvar(request("ino"),10)
if iCurrpage ="" then iCurrpage =1
iPageSize = 20
	
	Set clsPayForm = new CPayForm
		clsPayForm.Fpart_sn = part_sn
		clsPayForm.Fpatternname = sRectPatternName
		clsPayForm.FPageSize	= iPageSize
		clsPayForm.FCurrPage	=iCurrpage
		arrList = clsPayForm.fnGetPayPatternList
		iTotCnt = clsPayForm.FTotCnt
	Set clsPayForm = nothing 
		
	iTotalPage 	=  int((iTotCnt-1)/iPageSize) +1  '전체 페이지 수
%>
<html>
<head>
<title>계약정보 등록</title>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<link rel="stylesheet" href="/bct.css" type="text/css">
<script language="javascript">
<!--
	//신규등록
	function jsContsReg(patternSeq){
	location.href="pop_payform_pattern_reg.asp?sEN=<%=sEmpno%>&iPS="+patternSeq+"&iDPS=<%=iDefaultPaySeq%>&ino=<%=ino%>";
	}
	
	//패턴적용
	function jsSetPattern(patternSeq){
		opener.location.href="pop_payform.asp?sEN=<%=sEmpno%>&iPS="+patternSeq+"&iDPS=<%=iDefaultPaySeq%>&ino=<%=ino%>";
		self.close();
	}
	
	// 페이지 이동
	function jsGoPage(pg)
	{
		document.frm.iCP.value=pg;
		document.frm.submit();
	}
//-->
</script>
</head>
<body leftmargin="10" topmargin="10">
<table width="100%" border="0" cellpadding="5" cellspacing="0" class="a">
<tr>
	<td><strong>계약직사원 계약정보 패턴</strong><br><hr width="100%"></td>
</tr>
<tr>
	<td>
		<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">	
		<form name="frmSearch" method="post" action="pop_payform_pattern.asp">
		<input type="hidden" name="sEN" value="<%=sEmpno%>">
		<input type="hidden" name="iDPS" value="<%=iDefaultPaySeq%>">
			<input type="hidden" name="ino" value="<%=ino%>">
		<tr align="center" bgcolor="#FFFFFF" >
			<td rowspan="2" width="50" height="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
			<td align="left">&nbsp;&nbsp;&nbsp;부서: <%=printPartOption("part_sn", part_sn)%> &nbsp; &nbsp;
			패턴명: <input type="text" name="sRPN" value="<%=sRectPatternName%>" size="20" maxlength="60"></td>
			<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="검색" onClick="javascript:document.frmSearch.submit();">
			</td>
		</tr>
		</form>
		</table>
	</td>
</tr>
<tr>
	<td align="right"><input type="button" class="button" value="신규등록" onClick="jsContsReg('');"></td>
</tr>
<tr>
	<td>총: <%=iTotCnt%>건
		<table width="100%" border="0" cellpadding="5" cellspacing="1" align="center" class="a" bgcolor=#BABABA> 
		<form name="frm" method="get" action="">	
		<input type="hidden" name="sEN" value="<%=sEmpno%>">	
		<input type="hidden" name="iCP" value="">
		<input type="hidden" name="iDPS" value="<%=iDefaultPaySeq%>">
		<input type="hidden" name="ino" value="<%=ino%>">
		<tr bgcolor="<%= adminColor("tabletop") %>" align="center">
			<td>패턴번호</td>
			<td>부서</td>
			<td>패턴명</td>
			<td>선택</td>
		</tr>
		<%IF isArray(arrList) THEN
			For intLoop =0 To UBound(arrList,2)
			%>
		<tr align="center" bgcolor="#FFFFFF">
			<td><a href="javascript:jsContsReg(<%=arrList(0,intLoop)%>);"><%=arrList(0,intLoop)%></a></td>
			<td><a href="javascript:jsContsReg(<%=arrList(0,intLoop)%>);"><%=arrList(2,intLoop)%></a></td>
			<td><a href="javascript:jsContsReg(<%=arrList(0,intLoop)%>);"><%=arrList(3,intLoop)%></a></td>
			<td><input type="button" class="button" value="선택" onClick="jsSetPattern(<%=arrList(0,intLoop)%>);"></td>
		</tr>	
		<%	Next 
		ELSE
		%>
		<tr align="center" bgcolor="#FFFFFF">
			<td colspan="4">등록된 내용이 없습니다.</td>
		</tr>
		<%
		END IF%>
		</table>
	</td>
</tr>
<!-- 페이지 시작 -->
<%
Dim iStartPage,iEndPage,iX,iPerCnt
iPerCnt = 10

iStartPage = (Int((iCurrpage-1)/iPerCnt)*iPerCnt) + 1

If (iCurrpage mod iPerCnt) = 0 Then
	iEndPage = iCurrpage
Else
	iEndPage = iStartPage + (iPerCnt-1)
End If
%>
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15" align="center">
			<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a"  >
			    <tr valign="bottom" height="25">        
			        <td valign="bottom" align="center">
			         <% if (iStartPage-1 )> 0 then %><a href="javascript:jsGoPage(<%= iStartPage-1 %>)" onfocus="this.blur();">[pre]</a>
					<% else %>[pre]<% end if %>
			        <%
						for ix = iStartPage  to iEndPage
							if (ix > iTotalPage) then Exit for
							if Cint(ix) = Cint(iCurrpage) then
					%>
						<a href="javascript:jsGoPage(<%= ix %>)" class="menu_link3" onfocus="this.blur();"><font color="00abdf"><strong>[<%=ix%>]</strong></font></a>
					<%		else %>
						<a href="javascript:jsGoPage(<%= ix %>)" class="menu_link3" onfocus="this.blur();">[<%=ix%>]</a>
					<%
							end if
						next
					%>
			    	<% if Cint(iTotalPage) > Cint(iEndPage)  then %><a href="javascript:jsGoPage(<%= ix %>)" onfocus="this.blur();">[next]</a>
					<% else %>[next]<% end if %>
			        </td>        
			    </tr>        
			</table>
		</td>
	</tr>
	</form>
</table>
</body>
</html>
<!-- #include virtual="/lib/db/dbclose.asp" -->	
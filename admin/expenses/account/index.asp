<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 운영비 계정과목 리스트
' History : 2011.05.30 정윤정  생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"--> 
<!-- #include virtual="/lib/classes/expenses/OpExpArapCls.asp"-->
<%
Dim clsAccount, arrList, intLoop  	
Dim sarap_nm
Dim iTotCnt,iPageSize, iTotalPage,iCurrPage
 
	iPageSize = 20
	iCurrPage = requestCheckvar(Request("iCP"),10)
	if iCurrPage="" then iCurrPage=1
		 
 	sarap_nm =  requestCheckvar(Request("sAN"),30)
 	
Set clsAccount = new COpExpAccount
	clsAccount.Farap_nm  	= sarap_nm 
	clsAccount.FCurrPage 	= iCurrPage
	clsAccount.FPageSize 	= iPageSize
	arrList = clsAccount.fnGetOpExpAccountList 	
	iTotCnt = clsAccount.FTotCnt
Set clsAccount = nothing
	
	iTotalPage 	=  int((iTotCnt-1)/iPageSize) +1  '전체 페이지 수
%> 
 
<script language="javascript">
<!--
// 페이지 이동
function jsGoPage(iCP)
	{
		document.frm.iCP.value=iCP;
		document.frm.submit();
	}
	   
//새로등록
function jsNewReg(){
	var winC = window.open("popAccount.asp","popC","width=800, height=800, resizable=yes, scrollbars=yes");
	winC.focus();
}  

//삭제
function jsDelete(iOEA){
	if(confirm("삭제하시겠습니까?")){
		document.frmDel.hidOEA.value = iOEA;
		document.frmDel.submit();
	}
}

//수정
function jsMod(iOEA, iValue,strMsg){
	if(confirm(strMsg)){
		document.frmMod.hidOEA.value = iOEA;
		document.frmMod.hidInOut.value = iValue;
		document.frmMod.submit();
	}else{ 
	window.location.reload();
}
}
//-->
</script>
<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a"> 
<form name="frmDel" method="post" action="procAccount.asp">
<input type="hidden" name="hidM" value="D">
<input type="hidden" name="hidOEA" value="">
<input type="hidden" name="menupos" value="<%=menupos%>">
</form>
<form name="frmMod" method="post" action="procAccount.asp">
<input type="hidden" name="hidM" value="U">
<input type="hidden" name="hidOEA" value="">
<input type="hidden" name="hidInOut" value="">
<input type="hidden" name="menupos" value="<%=menupos%>">
</form>
<tr>
	<td>
		<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
			<form name="frm" method="get" action="">
			<input type="hidden" name="menupos" value="<%= menupos %>">
			<input type="hidden" name="iCP" value="">
			<tr align="center" bgcolor="#FFFFFF" >
				<td  width="100" height="50" bgcolor="<%= adminColor("gray") %>">검색 조건</td>
				<td align="left">
					계정과목: 
					 <input type="text" name="sAN" size="20" maxlenght="30" value="<%=sarap_nm%>">
				</td>
				<td  width="50" bgcolor="<%= adminColor("gray") %>">
					<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
				</td>
			</tr>
			</form>
		</table>
	</td>
</tr> 
<!-- #include virtual="/lib/db/dbclose.asp" --> 
<tr>
	<td><input type="button" class="button" value="수지항목 추가" onClick="jsNewReg();"></td>
</tr>
<tr>
	<td>
		<!-- 상단 띠 시작 -->
		<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
			<tr height="25" bgcolor="FFFFFF">
				<td colspan="15">
					검색결과 : <b><%=iTotCnt%></b> &nbsp;
					페이지 : <b><%= iCurrPage %> / <%=iTotalPage%></b>
				</td>
			</tr> 
			<tr align="center" bgcolor="<%= adminColor("tabletop") %>"> 
				<td>운영비코드</td>  
				<td>수지항목</td>  
				<td>연결계정과목</td> 
				<td>erpcode</td>  
				<td>사용/지급 금액구분</td>
				<td>처리</td>
			</tr>
			<%  
			IF isArray(arrList) THEN
				For intLoop = 0 To UBound(arrList,2) 
				%>
			<tr height=30 bgcolor="#FFFFFF">	
				<td align="center" > <%=arrList(0,intLoop)%></td>
				<td>[<%=arrList(1,intLoop)%>] <%=arrList(2,intLoop)%></td>
				<td>[<%=arrList(3,intLoop)%>] <%=arrList(4,intLoop)%></td>			
				<td align="center"><%=arrList(5,intLoop)%></td>	 
				<td align="center"><input type="radio" name="rdoInOut<%=arrList(0,intLoop)%>" value="1" <%IF arrList(6,intLoop) THEN %>checked<%END IF%> onClick="jsMod('<%=arrList(0,intLoop)%>',1,'사용금액으로 설정하시겠습니까?');">사용금액 <input type="radio" name="rdoInOut<%=arrList(0,intLoop)%>" value="0" <%IF not arrList(6,intLoop) THEN %>checked<%END IF%> onClick="jsMod('<%=arrList(0,intLoop)%>',0,'지급금액으로 설정하시겠습니까?');">지급금액</td>	 
				<td align="center"><input type="button" class="button" value="삭제" onClick="jsDelete('<%=arrList(0,intLoop)%>');"> </td>
			</tr>
		<%	Next
			ELSE%>
			<tr height=5 align="center" bgcolor="#FFFFFF">				
				<td colspan="6">등록된 내용이 없습니다.</td>	
			</tr>
			<%END IF%>
		</table>	
	</td> 
</tr> 	
<!-- 페이지 시작 -->
		<%
		Dim iStartPage,iEndPage,iX,iPerCnt
		iPerCnt = 10
		
		iStartPage = (Int((iCurrPage-1)/iPerCnt)*iPerCnt) + 1
		
		If (iCurrPage mod iPerCnt) = 0 Then
			iEndPage = iCurrPage
		Else
			iEndPage = iStartPage + (iPerCnt-1)
		End If
		%>
			<tr height="25" >
				<td colspan="15" align="center">
					<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
					    <tr valign="bottom" height="25">        
					        <td valign="bottom" align="center">
					         <% if (iStartPage-1 )> 0 then %><a href="javascript:jsGoPage(<%= iStartPage-1 %>)" onfocus="this.blur();">[pre]</a>
							<% else %>[pre]<% end if %>
					        <%
								for ix = iStartPage  to iEndPage
									if (ix > iTotalPage) then Exit for
									if Cint(ix) = Cint(iCurrPage) then
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
</table>
<!-- 페이지 끝 -->
</body>
</html>
 



	
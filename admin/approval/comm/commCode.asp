<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 공통코드관리 리스트
' History : 2011.03.09 정윤정  생성
'			2022.07.11 한용민 수정(isms취약점보안조치, 표준코드로변경)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/approval/commCls.asp"-->
<%
Dim clsComm, arrList, intLoop 
Dim iparentkey 
Dim iTotCnt,iPageSize, iTotalPage,iCurrPage
 
	iPageSize = 20
	iCurrPage = requestCheckvar(Request("iCP"),10)
	if iCurrPage="" then iCurrPage=1
	
	iparentkey = requestCheckvar(Request("selPK"),10)  
	if iparentkey = "" then iparentkey = 0 
Set clsComm = new CcommCode
	clsComm.Fparentkey 	= iparentkey 
	clsComm.FCurrPage 	= iCurrPage
	clsComm.FPageSize 	= iPageSize
	arrList = clsComm.fnGetCommCDList 	
	iTotCnt = clsComm.FTotCnt
	
	iTotalPage 	=  int((iTotCnt-1)/iPageSize) +1  '전체 페이지 수
%> 
 
<script type='text/javascript'>
<!--
// 페이지 이동
function jsGoPage(iCP)
	{
		document.frm.iCP.value=iCP;
		document.frm.submit();
	}
	  
function jsChangeGroup(){
	document.frm.submit();
}
	  
//새로등록
function jsNewReg(){
	var winC = window.open("popCommCodeConts.asp","popC","width=1200, height=768, resizable=yes, scrollbars=yes");
	winC.focus();
} 
//수정
function jsModReg(commCD){
	var winC = window.open("popCommCodeConts.asp?icc="+commCD,"popC","width=1200, height=768, resizable=yes, scrollbars=yes");
	winC.focus();
}

//-->
</script>
<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a"> 
<tr>
	<td>
		<form name="frm" method="get" action="">
		<input type="hidden" name="menupos" value="<%= menupos %>">
		<input type="hidden" name="iCP" value="">
		<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
			<tr align="center" bgcolor="#FFFFFF" >
				<td rowspan="2" width="100" height="50" bgcolor="<%= adminColor("gray") %>">검색 조건</td>
				<td align="left">
					그룹명  :
					<select name="selPK" onChange="jsChangeGroup();">
					<option value="0">--그룹--</option> 
					<%clsComm.FRectParentKey = iparentkey
					clsComm.sbOptCommCDGroup%>
					</select> 
				</td> 
				<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
					<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
				</td>
			</tr>
		</table>
		</form>
	</td>
</tr>
<%Set clsComm = nothing %>
<!-- #include virtual="/lib/db/dbclose.asp" --> 
<tr>
	<td><input type="button" class="button" value="신규등록" onClick="jsNewReg();"></td>
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
				<td>그룹명</td>
				<td>IDX</td>
				<td>추가코드</td>
				<td>코드명</td> 
				<td>설명</td> 	   
				<td>표시순서</td> 	
			</tr>
			<%  
			IF isArray(arrList) THEN
				For intLoop = 0 To UBound(arrList,2) 
				%>
			<tr height=30 align="center" bgcolor="#FFFFFF">	
				<td><a href="javascript:jsModReg(<%=arrList(0,intLoop)%>);"><%=arrList(5,intLoop)%></td>
				<td><a href="javascript:jsModReg(<%=arrList(0,intLoop)%>);"><%=arrList(0,intLoop)%></td>			
				<td><a href="javascript:jsModReg(<%=arrList(0,intLoop)%>);"><%=arrList(3,intLoop)%></a></td>	
				<td><a href="javascript:jsModReg(<%=arrList(0,intLoop)%>);"><%= ReplaceBracket(arrList(1,intLoop)) %></td>	
				<td><a href="javascript:jsModReg(<%=arrList(0,intLoop)%>);"><%=arrList(2,intLoop)%></td>  
				<td><a href="javascript:jsModReg(<%=arrList(0,intLoop)%>);"><%=arrList(6,intLoop)%></td>  
			</tr>
		<%	Next
			ELSE%>
			<tr height=5 align="center" bgcolor="#FFFFFF">				
				<td colspan="12">등록된 내용이 없습니다.</td>	
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
 



	
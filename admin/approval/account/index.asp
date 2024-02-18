<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 계정과목 내용  리스트
' History : 2011.03.09 정윤정  생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/approval/commCls.asp"-->
<!-- #include virtual="/lib/classes/approval/accountCls.asp"-->
<%
Dim clsComm, clsAccount, arrList, intLoop 
Dim iaccountkind, iedmsidx, saccountname, iparentkey 
Dim iTotCnt,iPageSize, iTotalPage,iCurrPage
 
	iPageSize = 20
	iCurrPage = requestCheckvar(Request("iCP"),10)
	if iCurrPage="" then iCurrPage=1
		
	iaccountkind =  requestCheckvar(Request("selAK"),10)
 	saccountname =  requestCheckvar(Request("sAN"),30)
 	
Set clsAccount = new CAccount
	clsAccount.Faccountkind 	= iaccountkind 
	clsAccount.Faccountname 	= saccountname 
	clsAccount.FCurrPage 	= iCurrPage
	clsAccount.FPageSize 	= iPageSize
	arrList = clsAccount.fnGetAccountList 	
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
	var winC = window.open("popAccountConts.asp","popC","width=600, height=400, resizable=yes, scrollbars=yes");
	winC.focus();
} 
//수정
function jsModReg(accountidx){
	var winC = window.open("popAccountConts.asp?iaidx="+accountidx,"popC","width=600, height=400, resizable=yes, scrollbars=yes");
	winC.focus();
}

//-->
</script>
<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a"> 
<tr>
	<td>
		<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
			<form name="frm" method="get" action="">
			<input type="hidden" name="menupos" value="<%= menupos %>">
			<input type="hidden" name="iCP" value="">
			<tr align="center" bgcolor="#FFFFFF" >
				<td rowspan="2" width="100" height="50" bgcolor="<%= adminColor("gray") %>">검색 조건</td>
				<td align="left">
					계정과목:
					<select name="selAK" >
					<option value="0">--전체--</option> 
					<%
					set clsComm = new CcommCode
					clsComm.Fparentkey = 1
					clsComm.Fcomm_cd = iaccountkind
					clsComm.sbOptCommCD
					Set clsComm = nothing
					%>
					</select>  
					&nbsp;&nbsp;
					계정과목 내용:
					 <input type="text" name="sAN" size="20" maxlenght="30" value="<%=saccountname%>">
				</td>
				<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
					<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
				</td>
			</tr>
			</form>
		</table>
	</td>
</tr> 
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
				<td>IDX</td>
				<td>계정과목내용</td> 
				<td>계정과목</td>
				<td>문서명</td>  
			</tr>
			<%  
			IF isArray(arrList) THEN
				For intLoop = 0 To UBound(arrList,2) 
				%>
			<tr height=30 align="center" bgcolor="#FFFFFF">	
				<td><a href="javascript:jsModReg(<%=arrList(0,intLoop)%>);"><%=arrList(0,intLoop)%></td>
				<td><a href="javascript:jsModReg(<%=arrList(0,intLoop)%>);"><%=arrList(3,intLoop)%></td>			
				<td><a href="javascript:jsModReg(<%=arrList(0,intLoop)%>);"><%=arrList(5,intLoop)%></a></td>	
				<td><a href="javascript:jsModReg(<%=arrList(0,intLoop)%>);"><%=arrList(7,intLoop)%></td>	 
			</tr>
		<%	Next
			ELSE%>
			<tr height=5 align="center" bgcolor="#FFFFFF">				
				<td colspan="4">등록된 내용이 없습니다.</td>	
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
 



	
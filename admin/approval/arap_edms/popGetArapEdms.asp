<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 수지항목 문서연동  리스트
' History : 2011.11.15  정윤정  생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/approval/araplinkedmsCls.asp"--> 
<%
Dim clsALE, arrList, intLoop 
Dim  sARAPNM,sedmsNM,iedmsIdx
Dim iTotCnt,iPageSize, iTotalPage,iCurrPage 
		
	sARAPNM =  requestCheckvar(Request("sANM"),50)
 	sedmsNM =  requestCheckvar(Request("sENM"),60)
  iedmsIdx =  requestCheckvar(Request("ieidx"),10)
  iPageSize = 20
	iCurrPage = requestCheckvar(Request("iCP"),10)
	if iCurrPage="" then iCurrPage=1 
	if iedmsIdx = "" then iedmsIdx = 0
		
Set clsALE = new CArapLinkEdms
	clsALE.FARAP_NM 	= sARAPNM 
	clsALE.FEdmsName 	= sedmsNM  
	clsALE.FedmsIdx		= iedmsIdx
	clsALE.FCurrPage	= iCurrPage
	clsALE.FPageSize	= iPageSize
	arrList = clsALE.fnGetEappArapLinkEdmsList
	iTotCnt = clsALE.FtotCnt 	 
Set clsALE = nothing 
 iTotalPage 	=  int((iTotCnt-1)/iPageSize) +1  '전체 페이지 수
 IF isArray(arrList) and iedmsIdx > 0 and  sedmsNM = "" THEN sedmsNM = arrList(5,0) 
%> 
 
<script language="javascript">
<!--
// 페이지 이동
function jsGoPage(iCP)
	{
		document.frm.iCP.value=iCP;
		document.frm.submit();
	} 
	
	
	function jsSelectEApp(iaidx,ieidx){
	opener.location.href= "/admin/approval/eapp/regeapp.asp?iAidx="+iaidx+"&ieidx="+ieidx; 
	self.close();
	}  
//-->
</script>
<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a"> 
<tr>
	<td><strong>수지항목 선택 - [ 문서명:<%=sedmsNM%> ]</strong><br><hr width="100%"></td>
</tr>
<tr>
	<td>
		<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
			<form name="frm" method="get" action=""> 
			<input type="hidden" name="iCP" value="">
			<input type="hidden" name="ieidx" value="<%=iedmsIdx%>">
			<input type="hidden" name="sENM" value="<%=sedmsNM%>">
			<tr align="center" bgcolor="#FFFFFF" >
				<td rowspan="2" width="100" height="50" bgcolor="<%= adminColor("gray") %>">검색 조건</td>
				<td align="left">
					수지항목: <input type="text" name="sANM" value="<%=sARAPNM%>" size="20"> 
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
	<td>
		<!-- 상단 띠 시작 -->
		<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>"> 
			<tr align="center" bgcolor="<%= adminColor("tabletop") %>"> 
				<td>IDX</td> 
				<td>수지항목</td> 
				<td>연결계정과목</td> 
			</tr>
			<%  
			IF isArray(arrList) THEN
				For intLoop = 0 To UBound(arrList,2) 
				%>
			<tr height=30 align="center" bgcolor="#FFFFFF">	
				<td><a href="javascript:jsSelectEApp('<%=arrList(1,intLoop)%>',<%=arrList(2,intLoop)%>);"><%=arrList(0,intLoop)%></a></td> 
				<td><a href="javascript:jsSelectEApp('<%=arrList(1,intLoop)%>',<%=arrList(2,intLoop)%>);"><%=arrList(3,intLoop)%></a></td>	
				<td><a href="javascript:jsSelectEApp('<%=arrList(1,intLoop)%>',<%=arrList(2,intLoop)%>);"><%=arrList(4,intLoop)%></a></td>	  
			</tr>
		<%	Next
			ELSE%>
			<tr height=5 align="center" bgcolor="#FFFFFF">				
				<td colspan="3">등록된 내용이 없습니다.</td>	
			</tr>
			<%END IF%>
		</table>	
	</td> 
</tr> 
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
		<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" >
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
</body>
</html>
 



	
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
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->  
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/linkedERP/arapCls.asp"-->
<!-- #include virtual="/lib/classes/approval/araplinkedmsCls.asp"--> 
<%
Dim clsALE, arrList, intLoop 
Dim sARAP_GB,sCASH_FLOW,sACC
Dim sARAPNM,sedmsNM, sMatch
Dim iTotCnt,iPageSize, iTotalPage,iCurrPage
		
	sARAP_GB = requestCheckvar(Request("rdoGB"),3)  
	sCASH_FLOW = requestCheckvar(Request("selFlow"),3)  	
	sARAPNM =  requestCheckvar(Request("sANM"),50)
	sACC =  requestCheckvar(Request("sAC"),50) 
 	sedmsNM =  requestCheckvar(Request("sENM"),60)
  sMatch=  requestCheckvar(Request("rdoM"),1) 
  iPageSize = 20
	iCurrPage = requestCheckvar(Request("iCP"),10)
	if iCurrPage="" then iCurrPage=1 
  if sMatch = "" then sMatch = "A"
   
 
Set clsALE = new CArapLinkEdms
	clsALE.FARAP_GB		= sARAP_GB
	clsALE.FCASH_FLOW	=	sCASH_FLOW
	clsALE.FACC				= sACC
	clsALE.FARAP_NM 	= sARAPNM 
	clsALE.FEdmsName 	= sedmsNM  
	clsALE.Fmatch 		= sMatch 
	clsALE.FCurrPage	= iCurrPage
	clsALE.FPageSize	= iPageSize
	arrList = clsALE.fnGetArapLinkEdmsList 	 
	iTotCnt = clsALE.FtotCnt 
Set clsALE = nothing
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
	   
 
//수정
function jsModReg(dAc){
	var winC = window.open("popArapEdms.asp?dAc="+dAc,"popC","width=600, height=400, resizable=yes, scrollbars=yes");
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
			<tr bgcolor="#FFFFFF" >
				<td rowspan="2" width="100" height="50" bgcolor="<%= adminColor("gray") %>">검색 조건</td>
				<td>
					 구분:
						<input type="radio" name="rdoGB" value=""<%IF sARAP_GB="" THEN%>checked<%END IF%>>전체
						<input type="radio" name="rdoGB" value="1" <%IF sARAP_GB="1" THEN%>checked<%END IF%>>수입
						<input type="radio" name="rdoGB" value="2" <%IF sARAP_GB="2" THEN%>checked<%END IF%>>지출
						&nbsp; &nbsp; &nbsp;
						분류:
						<select name="selFlow">
							<option value="">전체</option>
							<option value="001"  <%IF sCASH_FLOW="001" THEN%>selected<%END IF%>>영업</option>
							<option value="002"  <%IF sCASH_FLOW="002" THEN%>selected<%END IF%>>투자</option>
							<option value="003"  <%IF sCASH_FLOW="003" THEN%>selected<%END IF%>>재무</option>
						</select>
					&nbsp; &nbsp; &nbsp;
					문서매칭: <input type="radio" name="rdoM" value="A" <%IF sMatch="A" THEN%>checked<%END IF%>>전체
					<input type="radio" name="rdoM" value="Y" <%IF sMatch="Y" THEN%>checked<%END IF%>>매칭
					<input type="radio" name="rdoM" value="N" <%IF sMatch="N" THEN%>checked<%END IF%>>미매칭
				</td>
				<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
					<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
				</td>
			</tr>
			<tr  bgcolor="#FFFFFF" >
				<td align="left">
					수지항목: <input type="text" name="sANM" value="<%=sARAPNM%>" size="20"> 
					&nbsp;&nbsp;
					연결계정과목: <input type="text" name="sAC" value="<%=sACC%>" size="20">
					&nbsp;&nbsp;
					문서명: <input type="text" name="sENM" value="<%=sedmsNM%>" size="20"> 
				</td> 
			</tr>
			</form>
		</table>
	</td>
</tr> 
<!-- #include virtual="/lib/db/dbclose.asp" --> 
<%IF C_MngPart OR C_ADMIN_AUTH or C_PSMngPart THEN%>
<script language="javascript">
	function jsGetErp(){
		location.href = "/admin/linkedERP/arap/procGetErp.asp";
	}
</script> 
<tr>
	<td><input type="button" class="button" value="ERP목록수신" onClick="jsGetErp();"></td>
</tr>
<%END IF%>
<tr>
	<td>총:<%=iTotCnt%> 
		<!-- 상단 띠 시작 -->
		<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>"> 
			<tr align="center" bgcolor="<%= adminColor("tabletop") %>"> 
				<td>수지항목코드</td>
				<td>구분</td>  
				<td>분류</td>  
				<td>수지항목</td>  
				<td>연결계정과목</td>  
				<td>매입/매출거래종류</td>  
				<td>문서명</td>  
				<td>처리</td>  
			</tr>
			<%  
			IF isArray(arrList) THEN
				For intLoop = 0 To UBound(arrList,2) 
				%>
			<tr height=30 align="center" bgcolor="#FFFFFF">	
				<td><%=arrList(0,intLoop)%></td>
				<td><%=fnGetARAP_GB(arrList(1,intLoop))%></td> 
		 		<td><%=fnGetARAP_Cash(arrList(3,intLoop))%></td> 
				<td align="left"><%=arrList(2,intLoop)%><%IF arrList(11,intLoop) ="N" or arrList(12,intLoop)="Y" THEN%><font color="red"> [삭제 또는 비활성 항목]</font><%END IF%></td>			
				<td align="left">[<%=arrList(9,intLoop)%>] <%=arrList(5,intLoop)%><%IF arrList(13,intLoop) ="N" or arrList(14,intLoop)="Y" THEN%><font color="red"> [삭제 또는 비활성 항목]</font><%END IF%></a></td>	
				<td><%=arrList(7,intLoop)%></td>	  
				<td align="left"><%IF arrList(18,intLoop) <> "" THEN%>[<%=arrList(18,intLoop)%>] <%END IF%><%=arrList(16,intLoop)%><%IF not arrList(17,intLoop) THEN%><font color="red">[삭제문서]</font><%END IF%></td>	  
					<td><input type="button" value="<%IF arrList(15,intLoop) <> "" THEN%>수정<%ELSE%>등록<%END IF%>" class="button" onClick="jsModReg('<%=arrList(0,intLoop)%>')"></td>
			</tr>
		<%	Next
			ELSE%>
			<tr height=5 align="center" bgcolor="#FFFFFF">				
				<td colspan="5">등록된 내용이 없습니다.</td>	
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
 



	
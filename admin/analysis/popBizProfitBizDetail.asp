<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbiTmsOpen.asp" -->
<!-- #include virtual="/lib/db/dbiTMSHelper.asp"-->  
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/analysis/bizMonthProfitCls.asp"--> 
<!-- #include virtual="/lib/classes/linkedERP/bizSectionCls.asp"-->
<%
Dim intY, intM, sYear, sMonth
sYear = requestCheckvar(request("selY"),4)  
sMonth = requestCheckvar(request("selM"),2)   
IF sYear ="" THEN sYear = year(date)
IF sMonth ="" THEN sMonth = month(date) 
  
Dim bizSecCd : bizSecCd=requestCheckvar(request("bizSecCd"),16)
Dim accusecd : accusecd=requestCheckvar(request("accusecd"),16)     
  

''사업부문
Dim clsBS, arrBizList, intL,arrAllBizList
Dim intB, intC,oldbiz ,sumM,sumTot
Set clsBS = new CBizSection    
	clsBS.FYYYYMM = sYear&"-"& format00(2,sMonth) 
	arrBizList = clsBS.fnGetBizMonthUserBizList  
	clsBS.FUSE_YN = "Y"  
	clsBS.FOnlySub = "Y"  
	arrAllBizList = clsBS.fnGetBizSectionList 
Set clsBS = nothing
  
Dim clsBP, arrList, intLoop
Set clsBP = new CBizProfit
	clsBP.FYYYYMM =  sYear&"-"& format00(2,sMonth) 
	clsBP.Faccusecd = accusecd
	clsBP.FBizsection_Cd = bizSecCd 
	arrList = clsBP.fnGetBizMonthProfitBizDetail  
Set clsBP = nothing
  


%>

<script language='javascript'>
//검색
function jsSearch(){
	document.frm.submit();
}
</script>
<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a">
	<tr>
	<td>
		<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
			<form name="frm" method="get" action="">
			<input type="hidden" name="page" value="">
			<tr align="center" bgcolor="#FFFFFF" >
				<td  width="100" height="50" bgcolor="<%= adminColor("gray") %>">검색 조건</td>
				<td align="left"> 
					날짜:
					<select name="selY" id="selY">
						<%For intY = Year(date) To 2012 STEP -1%>
						<option value="<%=intY%>" <%IF Cstr(sYear) = Cstr(intY) THEN%>selected<%END IF%>><%=intY%></option>
						<%Next%>
					</select>
						<select name="selM">
						<%For intM = 1 To 12%>
						<option value="<%=intM%>" <%IF Cstr(sMonth) = Cstr(intM) THEN%>selected<%END IF%>><%=intM%></option>
						<%Next%>
					</select> 
					 &nbsp;&nbsp;
						사업부문:
                    <select name="bizSecCd">
                    <option value="">--선택--</option>
                    <% 
                    IF isArray(arrAllBizList) THEN
                    For intLoop = 0 To UBound(arrAllBizList,2)	%>
                		<option value="<%=arrAllBizList(0,intLoop)%>" <%IF Cstr(bizSecCd) = Cstr(arrAllBizList(0,intLoop)) THEN%> selected <%END IF%>><%=arrAllBizList(1,intLoop)%></option>
                	<% Next 
                END IF
                	%>
                    </select>
                    &nbsp;&nbsp;
                    계정과목코드:
					<input type="text" name="accusecd" value="<%=accusecd%>" size="15">
				</td>
				<td   width="50" bgcolor="<%= adminColor("gray") %>">
					<input type="button" class="button_s" value="검색" onClick="jsSearch();">
				</td>
			</tr> 
			</form>
		</table>
	</td>
</tr>
</table>

<p>
<!-- 상단 띠 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>"> 
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	  <td rowspan="2">날짜</td>
		<td colspan="4">계정</td> 
		<td rowspan="2">구분</td>
		<td <%IF isArray(arrBizList) THEN%>colspan="<%=uBound(arrBizList,2)+1%>"<%END IF%>>지원부서</td>
		<td rowspan="2">합계</td>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>"> 
		<td>계정그룹</td>
		<td>계정분류</td>
		<td>계정코드</td> 
		<td>계정과목</td>  
	 <%IF isArray(arrBizList) THEN
	 		For intL = 0 To uBound(arrBizList,2)
	 	%>
	 	<td><%=arrBizList(1,intL)%></td>
	 <%	Next
	 END IF%> 
	</tr> 
	<%intC = 0
		sumM = 0
		sumTot =0
	IF isArray(arrList) THEN
	%>
		<tr bgcolor="#FFFFFF">
			<td align="center"><%=arrList(0,0)%></td>
			<td align="center"><%=arrList(2,0)%></td>
			<td align="center"><%=arrList(4,0)%></td>
			<td align="center"><%=arrList(6,0)%></td> 
			<td align="center"><%=arrList(7,0)%></td> 
			<td align="center"><%=arrList(9,0)%></td>		
	<%	
		For intLoop = 0 To UBound(arrList,2)  
				IF Cstr(oldbiz) <> Cstr(arrList(8,intLoop)) and intLoop > 0  THEN	 
		%>
			<%For intB= intC To (intL-1) %>
			<td align="right">0</td>
			<%		Next%>
			<td  align="right"><%=formatnumber(sumM,0)%></td> 
		</tr>		
		<tr bgcolor="#FFFFFF">		
			<td align="center"><%=arrList(0,intLoop)%></td>
			<td align="center"><%=arrList(2,intLoop)%></td>
			<td align="center"><%=arrList(4,intLoop)%></td>
			<td align="center"><%=arrList(6,intLoop)%></td> 
			<td align="center"><%=arrList(7,intLoop)%></td> 
			<td align="center"><%=arrList(9,intLoop)%></td> 
		<%	intC = 0	
				sumM = 0
			END IF%> 
		<%IF isArray(arrBizList) THEN
			For intB = intC To (intL-1) 
				IF arrBizList(0,intB) = arrList(10,intLoop) THEN
		%>
			<td align="right"><%=formatnumber(arrList(12,intLoop),0)%></td>
	<%			intC = intB+1 
					sumM = sumM + arrList(12,intLoop)
					sumTot = sumTot+ sumM
					Exit For  
				ELSE
	%>
			<td align="right">0</td>
	<%			
				END IF 
		Next
		END IF	  
		oldbiz = arrList(8,intLoop) 
	Next
	END IF
		For intB= intC To (intL-1) 
		%>
		<td align="right">0</td>
		<%	Next%>
		<td  align="right"><%=formatnumber(sumM,0)%></td>
	</tr>	
	<tr  bgcolor="#DDFFDD">
		<td colspan="6" align="center">합계</td>
		<td align="right" <%IF isArray(arrBizList) THEN%>colspan="<%=uBound(arrBizList,2)+2%>"<%END IF%>><%=formatnumber(sumTot,0)%></td>
	</tr>		 
</table>	 
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbiTmsClose.asp" -->
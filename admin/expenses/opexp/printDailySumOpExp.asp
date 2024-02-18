<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 운영비관리 상세   리스트
' History : 2011.06.03 정윤정  생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"--> 
<!-- #include virtual="/lib/classes/expenses/OpExpAccountCls.asp"-->
<!-- #include virtual="/lib/classes/expenses/OpExpPartCls.asp"-->
<!-- #include virtual="/lib/classes/expenses/OpExpCls.asp"--> 
<%
Dim clsPart,clsOpExp,arrPart, arrList, arrType, intLoop 
Dim clsAccount, arrAccount  
Dim dYear, dMonth, iPartTypeIdx, iOpExpPartIdx, iarap_cd
Dim intY, intM
Dim isearchType
Dim iOpExpIdx,dyyyymm, mLastMonthExp,mInExp,mOutExp,mTotExp,sOpExpPartName,sPartTypeName
Dim iAuthValue  
 	dYear			= requestCheckvar(Request("selY"),4)
 	dMonth			= requestCheckvar(Request("selM"),2)
 	isearchType		= requestCheckvar(Request("rdoST"),1) 
 	IF isearchType = "" THEN isearchType =1
 	IF isearchType = 1 THEN
 	iPartTypeIdx	= requestCheckvar(Request("selPT"),10) 
 	iOpExpPartIdx	= requestCheckvar(Request("selP"),10) 
 	ELSE
 	iarap_cd		= requestCheckvar(Request("selA"),10)
	END IF
 
	iOpExpIdx		= requestCheckvar(Request("hidOE"),10)
 	IF dYear = "" THEN dYear = year(date())
 	IF dMonth = "" THEN dMonth = month(date())	
 		dyyyymm =  dYear&"-"&Format00(2,dMonth) 
 	IF 	iPartTypeIdx = "" THEN iPartTypeIdx = 0
 	IF 	iOpExpPartIdx = "" THEN iOpExpPartIdx = 0
 	IF 	iarap_cd = "" THEN iarap_cd = 0
		
	iAuthValue = 0
	 				
 '운영비관리 팀 구분 리스트		
Set clsPart = new COpExpPart
	arrType = clsPart.fnGetOpExpPartTypeList 
	
	IF iPartTypeIdx > 0 THEN
	clsPart.FPartTypeidx 	= iPartTypeIdx  
	arrPart = clsPart.fnGetOpExppartAllList  
	
	clsPart.fnGetOpExpPartTypeData 
	sPartTypeName  =clsPart.FPartTypeName  
	END IF
Set clsPart = nothing
 
'운영비 리스트	
Set clsOpExp = new OpExp 
	clsOpExp.FYYYYMM 		=dyyyymm
	clsOpExp.FOpExpPartIdx 	= iOpExpPartIdx   
	clsOpExp.FOpExpIdx 	= iOpExpIdx   
	clsOpExp.fnGetOpExpMonthlyData
	iOpExpidx 	   =  clsOpExp.FOpExpidx 	  
	dyyyymm		   =  clsOpExp.Fyyyymm		 
	iOpExpPartIdx   =  clsOpExp.FOpExpPartIdx 
	mLastMonthExp   =  clsOpExp.FLastMonthExp 
	mInExp		   =  clsOpExp.FInExp		 
	mOutExp		   =  clsOpExp.FOutExp		 
	mTotExp 	    =  clsOpExp.FTotExp 	 
	sOpExpPartName  =  clsOpExp.FOpExpPartName 
	iPartTypeIdx	= clsOpExp.FPartTypeIdx
 
	clsOpExp.FYYYYMM 		= dyyyymm 
	clsOpExp.FpartTypeidx 	= iPartTypeIdx  
	clsOpExp.FOpExpPartIdx 	= iOpExpPartIdx  
	clsOpExp.Farap_cd 		= iarap_cd  
	arrList = clsOpExp.fnGetOpExpDailySumList 
Set clsOpExp = nothing	
 
%>  
<!-- #include virtual="/lib/db/dbclose.asp" --> 
<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a">  
<tr>
	<td>+   <%=dYear%>년 <%=dMonth%>월 운영비 내역 - <%=sPartTypeName%> > <%=sOpExpPartName%> </td>
</tr>   
<tr>
	<td><table width="100%" align="center" cellpadding="5" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
		<tr>
			<td align="center" bgcolor="<%= adminColor("tabletop") %>">전월잔액</td>
			<td align="center" bgcolor="<%= adminColor("tabletop") %>">지급액</td>
			<td align="center" bgcolor="<%= adminColor("tabletop") %>">사용액</td>
			<td align="center" bgcolor="<%= adminColor("tabletop") %>">당월잔액</td>
			
		</tr>
		<tr> 
			<td align="center" bgcolor="#FFFFFF"><%=formatnumber(mLastMonthExp,0)%></td>
			<td align="center" bgcolor="#FFFFFF"><%=formatnumber(mInExp,0)%></td>
			<td align="center" bgcolor="#FFFFFF"><%=formatnumber(mOutExp,0)%></td>
			<td align="center" bgcolor="#FFFFFF"><%=formatnumber(mTotExp,0)%></td>
		</tr>
		</table>
</td>
</tr> 
<tr>
	<td>
		<!-- 상단 띠 시작 -->
		<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>"> 
			<tr align="center" bgcolor="<%= adminColor("tabletop") %>">   
				<%IF iPartTypeIdx > 0 THEN%>
				<td>수지항목</td> 
				<%ELSE%>
				<td>운영비사용처</td>   
				<%END IF%> 
				<td>지급액</td>    
				<td>사용액</td>   
				<td>공급가액</td>   
				<td>부가세</td>  
				<td>건수</td>  
			</tr>
			<%    Dim sumIn, sumOut, sumSup, sumVat,sumCnt
			IF isArray(arrList) THEN
				sumIn = 0
				sumOut = 0
				sumSup = 0
				sumVat = 0
				sumCnt = 0
				For intLoop = 0 To UBound(arrList,2)  
			 %>  
			<tr height=30 align="center" bgcolor="#FFFFFF">	
				<td><%IF isearchType="2" THEN%><%=arrList(7,intLoop)%> > <%END IF%><%=arrList(6,intLoop)%></td>
				<td><%=formatnumber(arrList(0,intLoop),0)%></td>
				<td><%=formatnumber(arrList(1,intLoop),0)%></td>
				<td><%=formatnumber(arrList(2,intLoop),0)%></td>
				<td><%=formatnumber(arrList(3,intLoop),0)%></td>
				<td><%=formatnumber(arrList(4,intLoop),0)%></td> 
			</tr>	
			<%	sumIn = sumIn + arrList(0,intLoop)
				sumOut = sumOut + arrList(1,intLoop)
				sumSup = sumSup + arrList(2,intLoop)	
				sumVat = sumVat + arrList(3,intLoop)
				sumCnt = sumCnt + arrList(4,intLoop)
			Next  
			ELSE%>
			<tr height="30" align="center" bgcolor="#FFFFFF">				
				<td colspan="6">등록된 내용이 없습니다.</td>	
			</tr>
			<%END IF%>
			<tr height=30 align="center" bgcolor="<%=adminColor("sky")%>">	
				<td>총합</td>
				<td><%=formatnumber(sumIn,0)%></td>
				<td><%=formatnumber(sumOut,0)%></td>
				<td><%=formatnumber(sumSup,0)%></td>
				<td><%=formatnumber(sumVat,0)%></td>
				<td><%=formatnumber(sumCnt,0)%></td> 
			</tr>
		</table>	
	</td> 
</tr> 	 
</table> 
<script language="javascript">
<!--
 	document.body.onload=function(){window.print();} 
//-->
</script> 
</body>
</html> 



	
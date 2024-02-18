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
<!-- #include virtual="/lib/classes/expenses/OpExpCls.asp"-->
<!-- #include virtual="/lib/classes/expenses/OpExpPartCls.asp"-->
<!-- #include virtual="/lib/classes/expenses/OpExpArapCls.asp"-->
<!-- #include virtual="/lib/classes/approval/partMoneyCls.asp"-->
<%
Dim clsPart,clsOpExp, arrPart, arrList, arrType, intLoop, clsPartMoney
Dim clsAccount, arrAccount ,iarap_cd
Dim  arrUsePart ,sOpExpPartName, sPartTypeName
Dim dYear, dMonth, iPartTypeIdx, iOpExpPartIdx	,sbizsection_cd, sbizsection_nm
Dim intY, intM
Dim iTotCnt,iPageSize, iTotalPage,iCurrPage
Dim iAuthValue
 
	iPageSize = 20
	iCurrPage = requestCheckvar(Request("iCP"),10)
	if iCurrPage="" then iCurrPage=1
		 
 	iPartTypeIdx	= requestCheckvar(Request("selPT"),10)
 	iOpExpPartIdx	= requestCheckvar(Request("selP"),10)
 	dYear			=  requestCheckvar(Request("selY"),4)
 	dMonth			=  requestCheckvar(Request("selM"),2)
 	iarap_cd		= requestCheckvar(Request("selA"),10)
 	sbizsection_nm=requestCheckvar(Request("sBiznm"),10)
 	IF dYear = "" THEN dYear = year(date())
 	IF dMonth = "" THEN dMonth = month(date())	
 	iAuthValue = 0	
 	 
 '운영비관리 팀 구분 리스트		
Set clsPart = new COpExpPart
	arrType = clsPart.fnGetOpExpPartTypeList 
	IF iPartTypeIdx > 0 THEN
	clsPart.FPartTypeidx 	= iPartTypeIdx  
	arrPart = clsPart.fnGetOpExppartAllList  
	END IF
 
	 
	clsPart.FOpExpPartidx = iOpExpPartIdx
	clsPart.fnGetOpExpPartName
	sOpExpPartName =clsPart.FOpExpPartName
	sPartTypeName  =clsPart.FPartTypeName 
Set clsPart = nothing
 	
'운영비 리스트	
Set clsOpExp = new OpExp
	clsOpExp.FYYYYMM 	= dYear&"-"&Format00(2,dMonth)
	clsOpExp.FPartTypeIdx = iPartTypeIdx 
	clsOpExp.FOpExpPartIdx = iOpExpPartIdx 
	clsOpExp.Farap_cd = iarap_cd
	clsOpExp.Fbizsection_nm = sbizsection_nm
	clsOpExp.FCurrPage 	= iCurrPage
	clsOpExp.FPageSize 	= iPageSize
	arrList = clsOpExp.fnGetOpExpDailyList
	iTotCnt = clsOpExp.FTotCnt  
Set clsOpExp = nothing	  
%>  
<!-- #include virtual="/lib/db/dbclose.asp" -->   
<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a">  
<tr>
	<td>+  <%=dYear%>년	<%=dMonth%>월 운영비 상세내역 - <%=sPartTypeName%> > <%=sOpExpPartName%></td>
</tr> 
<tr>
	<td> 
				<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a"    border="1">  
					<tr align="center" bgcolor="<%= adminColor("tabletop") %>">  
						<td width="50">순번</td>
						<td width="50">날짜(일)</td>  
						<td>운영비사용처</td>   
						<td>수지항목</td>
						<td>업체명</td>  
						<td>적요(상세내역)</td>   
						<td>지급액</td>  
						<td>사용액</td>   
						<td>공급가액</td> 
						<td>부가세</td> 
						<td>승인번호</td> 
						<td>사용부서</td>    
					</tr>
					<%   Dim totInExp, totOutExp,sumInExp,sumOutExp, iNum, sumSupExp, sumVatExp, totSupExp, totVatExp
					totInExp = 0
					totOutExp = 0
					sumInExp=0
					sumOutExp=0
					sumSupExp=0
					sumVatExp=0
					totSupExp=0
					totVatExp=0
					iNum = 1
					IF isArray(arrList) THEN
						For intLoop = 0 To UBound(arrList,2)  
					 %>  
					<tr height=30 bgcolor="#FFFFFF">	 
						<td align="center"><%=iNum%></td>
						<td align="center"><%=day(arrList(1,intLoop))%></td>
						<td align="center"><%=arrList(12,intLoop)%> > <%=arrList(11,intLoop)%></td>
						<td align="center"><%=arrList(3,intLoop)%></td>
						<td><%=arrList(6,intLoop)%></td>
						<td><%=arrList(7,intLoop)%></td> 
						<td align="right"><%=formatnumber(arrList(4,intLoop),0)%></td>
						<td align="right"><%=formatnumber(arrList(5,intLoop),0)%></td> 
						<td align="right"><%=formatnumber(arrList(8,intLoop),0)%></td> 
						<td align="right"><%=formatnumber(arrList(9,intLoop),0)%></td> 
						<td align="center"><%=arrList(10,intLoop)%>&nbsp;</td> 
						<td align="center"><%=arrList(13,intLoop)%>&nbsp;</td>  
					</tr>	
					<%
					  totInExp = totInExp + arrList(4,intLoop)
					  totOutExp = totOutExp + arrList(5,intLoop)
					  totSupExp = totSupExp + arrList(8,intLoop)
					  totVatExp = totVatExp + arrList(9,intLoop)
					  	
					  sumInExp = sumInExp +  arrList(4,intLoop)
					  sumOutExp = sumOutExp +  arrList(5,intLoop)	 
					  sumSupExp = sumSupExp +  arrList(8,intLoop)
					  sumVatExp = sumVatExp +  arrList(9,intLoop)	
					  
					  iNum = iNum + 1
				IF intLoop  < UBound(arrList,2)  THEN
					IF Cstr(arrList(2,intLoop)) <> Cstr(arrList(2,intLoop+1)) THEN%>
				   <tr height=30 align="center" bgcolor="#FFFFFF"> 
				   	<td colspan="6"><b><%=arrList(3,intLoop)%></b></td>
				   	<td align="right"><b><%=formatnumber(sumInExp,0)%></b></td>
				   	<td align="right"><b><%=formatnumber(sumOutExp,0)%></b></td>
				   	<td align="right"><%=formatnumber(sumSupExp,0)%></td>
				   	<td align="right"><%=formatnumber(sumVatExp,0)%></td>
				    <td colspan="3">&nbsp;</td> 
				</tr>
				<%	sumInExp = 0
					sumOutExp = 0
					sumSupExp = 0
					sumVatExp = 0
					iNum = 1
					END IF
				END IF
					Next  %>
					<tr  height=30 align="center" bgcolor="#FFFFFF"> 
				   	<td colspan="6"><b><%=arrList(3,intLoop-1)%></b></td>
				   	<td align="right"><b><%=formatnumber(sumInExp,0)%></b></td>
				   	<td align="right"><b><%=formatnumber(sumOutExp,0)%></b></td>
				   	<td align="right"><%=formatnumber(sumSupExp,0)%></td>
				   	<td align="right"><%=formatnumber(sumVatExp,0)%></td>
				   	<td colspan="3">&nbsp;</td> 
					</tr>
					<%
					ELSE%>
					<tr height="30" align="center" bgcolor="#FFFFFF">				
						<td colspan="13">등록된 내용이 없습니다.</td>	
					</tr>
					<%END IF%>
				<%IF iOpExpPartIdx > 0 THEN	 %>
				 <tr  height=30 align="center" bgcolor="#DDDDFF"> 
				   	<td colspan="6">총합</td>
				   	<td align="right"><%=formatnumber(totInExp,0)%></td>
				   	<td align="right"><%=formatnumber(totOutExp,0)%></td>
				   	<td align="right"><%=formatnumber(totSupExp,0)%></td>
				   	<td align="right"><%=formatnumber(totVatExp,0)%></td>
				   	<td colspan="3">&nbsp;</td> 
				</tr>
			 <%END IF%>
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



	
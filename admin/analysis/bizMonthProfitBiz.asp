<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->  
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"--> 
<!-- #include virtual="/lib/classes/linkedERP/bizSectionCls.asp"-->
<!-- #include virtual="/lib/classes/analysis/bizMonthProfitCls.asp"-->
<%  
Dim bizSecCd : bizSecCd=requestCheckvar(request("bizSecCd"),16)
Dim accusecd  : accusecd=requestCheckvar(request("accusecd"),16)
Dim isINTrans  : isINTrans=requestCheckvar(request("isINTrans"),10) ''내부거래 
Dim intY, intM, sYear, sMonth
Dim icolType, irowType, sGRP_YN
Dim oldacc, sumM, intI, intC,oldbizType
Dim accType, oldgrp, oldup, sumgrp, sumup,intSum
Dim dSDate, dEDate
sYear = requestCheckvar(request("selY"),4)  
sMonth = requestCheckvar(request("selM"),2)  
icolType = requestCheckvar(request("rdoC"),1)  '보기옵션 가로(1:사업부별,2:팀별)
irowType = requestCheckvar(request("rdoR"),1)  '보기옵션 세로(1:계정그룹,2:계정분류,3:계정과목)

IF sYear ="" THEN sYear = year(date)
IF sMonth ="" THEN sMonth = month(date) 
	dSDate = dateserial(sYear,sMonth,"1")
 	dEDate = dateadd("d",-1,dateadd("m",1,dSDate))
 
IF Len(accusecd)=3 then accusecd=accusecd&"00"
IF icolType = "" THEN icolType =1  
IF irowType = "" THEN irowType =1
IF icolType ="1" THEN  
	sGRP_YN="Y"
ELSE
	sGRP_YN="N"
END IF	
	
''사업부문
Dim clsBS, arrBizList, intL 
Set clsBS = new CBizSection    
	clsBS.FYYYYMM = sYear&"-"& format00(2,sMonth) 
	arrBizList = clsBS.fnGetBizMonthBizList  
Set clsBS = nothing

'손익보고서리스트			
Dim clsBP, arrList, intLoop
Set clsBP = new CBizProfit
	clsBP.FYYYYMM =  sYear&"-"& format00(2,sMonth) 
	clsBP.Faccusecd = accusecd
	clsBP.FbizType = isINTrans 
	arrList = clsBP.fnGetBizMonthProfitBizList  
Set clsBP = nothing
 
 
'사업부(arrPBiz), 팀별(arrBiz) 각각 변수에 저장 후 변수 불러오기  
Dim arrPBiz(),arrBiz(),intB,intP, intChk, oldPCD, sBizcd,ichkNull,arrgrp(),arrup() 
	intB = 0
	intP = 0
	intChk = 0
 IF isArray(arrBizList) THEN 
		For intLoop = 0 To UBound(arrBizList,2)    
			IF oldPCD <> arrBizList(2,intLoop) THEN
				intP = intP + 1
				redim preserve arrPBiz(2,intP)
				arrPBiz(1,intP) =  arrBizList(3,intLoop)
				arrPBiz(2,intP) =  arrBizList(2,intLoop) 
				IF intP> 1 THEN 
				arrPBiz(0,intP-1) = intChk
				END IF
				intChk =0
			 END IF
			 
				intChk = intChk + 2
			
				redim preserve arrBiz(2,intB+2)
				arrBiz(1,intB) = arrBizList(1,intLoop)  
				arrBiz(0,intB) = arrBizList(0,intLoop)
				arrBiz(2,intB) = False
				arrBiz(1,intB+1) = arrBizList(1,intLoop) &"<br>지원" 
				arrBiz(0,intB+1) = arrBizList(0,intLoop)
				arrBiz(2,intB+1) = True
			IF intLoop =  UBound(arrBizList,2)   THEN
					arrPBiz(0,intP) = intChk
			END IF
				intB = intB + 2
			oldPCD  = arrBizList(2,intLoop) 
		Next 
	END IF 
%>
 <script type="text/javascript">
 	
 	function jsSearch(){
 		document.frm.submit();
 	}
 	
 	function jsFillCal(val1, val2){   
 		for(i=0;i<document.all.selY.length;i++){
	    if(document.all.selY.options[i].value == val1){ 
	    	document.all.selY.options[i].selected = true;
	    }  
   }
  
    if(document.all.selM.options[parseInt(val2)-1].value == val2){ 
    	document.all.selM.options[parseInt(val2)-1].selected = true;
    }
}


function showProfitBizDetail(bizSecCd,accusecd){ 
    var iURI = "popBizProfitBizDetail.asp?selY=<%=sYear%>&selM=<%=sMonth%>&bizSecCd="+bizSecCd+"&accusecd="+accusecd+"&isINTrans=<%=isINTrans%>";
    var popwin = window.open(iURI,'showProfitBizDetail','scrollbars=yes,resizable=yes,width=900,height=600');
    popwin.focus();
}

function jsUpdateMPBiz(){
		document.frmU.submit();
}
 </script>
 <form name="frmU" method="post" action="procBizMonthProfit.asp">
 	<input type="hidden" name="hidM" value="B">
	<input type="hidden" name="hidYM" value="<%= sYear&"-"& format00(2,sMonth) %>"> 
</form>	
<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a">
	<tr>
	<td>
		<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
			<form name="frm" method="get" action="bizmonthprofitBiz.asp">
			<input type="hidden" name="menupos" value="<%= menupos %>">
			<input type="hidden" name="page" value=""> 
			<tr align="center" bgcolor="#FFFFFF" >
				<td  width="100" height="30" bgcolor="<%= adminColor("gray") %>">검색 조건</td>
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
					<input type="button" value="전전달" class="button" onClick="jsFillCal('<%= Year(DateAdd("m",-2,now()))%>','<%= Month(DateAdd("m",-2,now()))%>')";>
					<input type="button" value="전달" class="button" onClick="jsFillCal('<%= Year(DateAdd("m",-1,now()))%>','<%= Month(DateAdd("m",-1,now()))%>')";>
					<input type="button" value="이번달" class="button" onClick="jsFillCal('<%= Year(DateAdd("m",0,now()))%>','<%= Month(DateAdd("m",0,now()))%>')";>
					
					&nbsp;&nbsp;
					<input type="checkbox" name="isINTrans" value="2" <%= ChkIIF(isINTrans="2","checked","") %> > 내부거래만
					&nbsp;&nbsp;
					  계정과목코드:
					<input type="text" name="accusecd" value="<%=accusecd%>" size="15"> 
				</td>
				<td    width="50" bgcolor="<%= adminColor("gray") %>">
					<input type="button" class="button_s" value="검색" onClick="jsSearch();">
				</td>
			</tr>  
			</form>
		</table>
	</td>
</tr>  
<tr>
	<td><input type="button" class="button" value="update" onClick="jsUpdateMPBiz();"> : 검색조건의 날짜에 해당하는 데이터만 업데이트 됩니다.</td>
</tr>
<tr>
	<td>
		<!-- 상단 띠 시작 -->
		<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">  
		<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
			<td colspan="5"  rowspan="2" >계정과목</td>  
			<td rowspan="2">NULL </td>
			<%IF isArray(arrPBiz) THEN
				For intLoop = 1 To intP%>
			<td colspan="<%=arrPBiz(0,intLoop)%>"><%=arrPBiz(1,intLoop)%></td> 
			<%Next
				END IF
			%>  
			<td rowspan="2" >합계</td>
		</tr>   
		<tr  bgcolor="<%= adminColor("tabletop") %>"  align="center">   
			<%IF isArray(arrBizList) THEN
			For intLoop=0 To intB-1 %> 
				<td <%IF arrBiz(2,intLoop)  THEN%>bgcolor="#F5EF80"<%END IF%>><%=arrBiz(1,intLoop)%></td> 
			<%Next
		END IF
			%> 
		</tr>  
	<% 	intI = 0	
	sumgrp = 0
	sumup = 0 
	IF isArray(arrBizList) THEN 
		redim arrgrp(intB+1)
		redim arrup(intB+1)
	else
		redim arrgrp(1)
		redim arrup(1)
	END IF
	IF isArray(arrList) THEN 
			For intLoop = 0 To UBound(arrList,2)
				ichkNull =0 
					accType =arrList(8,intLoop)	'계정과목 
				IF oldacc <> accType or oldbizType <> arrList(5,intLoop) THEN	'계정과목이 틀려지면 table줄바꿈   
					IF intLoop > 0 THEN
						For intC= intI To (intB-1)'
				%>
				<td align="right">0</td>
			<%		Next
					intI = 0	
			%>
				<td  align="right"><%=formatnumber(sumM,0)%></td>
			</tr>
			
			<%  
						if oldup <> arrList(7,intLoop) and oldup <> "" THEN
							sumup = 0 
				%>
				<tr bgcolor="#EEEEEE"> 
				<td colspan="5" align="center"><%=oldup%></td> 
					<%For intSum = 0 To intB 
						sumup = sumup + arrup(intSum)
				%>
				<td  align="right"><%=formatnumber(arrup(intSum),0)%></td> 
				<%Next%>
					<td  align="right"><%=formatnumber(sumup,0)%></td> 
				</tr>
				<%	 redim arrup(intB+1)
				END IF 
				%>	
				<% IF   oldgrp <> arrList(1,intLoop) THEN
						sumgrp = 0 
				%>
				<tr bgcolor="#DDFFDD"> 
				<td colspan="5" align="center" style="border-bottom:2px solid <%=adminColor("tablebg")%>;"><%=oldgrp%></td> 
				<%For intSum = 0 To intB 
						sumgrp = sumgrp + arrgrp(intSum)
				%>
				<td  align="right"  style="border-bottom:2px solid <%=adminColor("tablebg")%>;"><%=formatnumber(arrgrp(intSum),0)%></td> 
				<%Next%>
					<td  align="right"  style="border-bottom:2px solid <%=adminColor("tablebg")%>;"><%=formatnumber(sumgrp,0)%></td> 
				</tr>  
				<%	redim arrgrp(intB+1) 
				END IF
				%>	
				
			<% 	END IF%>
			<tr align="center" bgcolor="#FFFFFF">
				<!--계정내용--> 
				<td><%=arrList(0,intLoop)%></td>
				<td><%=arrList(1,intLoop)%></td>
				<td><%=arrList(8,intLoop)%></td>
				<td><%=arrList(9,intLoop)%></td>
				<td><%=arrList(10,intLoop)%><%IF arrList(5,intLoop) ="2" THEN%><font color="blue">(내부거래)</font><%END IF%></td>  
				<!--/계정내용-->
			<%	sumM = 0 
					IF isNull(arrList(2,intLoop)) or arrList(2,intLoop) ="" THEN	'사업부서 Null 일때
					sumM = sumM + arrList(4,intLoop)-arrList(3,intLoop)
					 arrgrp(0) = arrgrp(0) + arrList(4,intLoop)-arrList(3,intLoop)
					 	arrup(0) = arrup(0)+ arrList(4,intLoop)-arrList(3,intLoop)
			%>	
				<td  align="right"><%=formatnumber(arrList(4,intLoop)-arrList(3,intLoop),0)%></td>		
		<%   	ichkNull = 1
				ELSE %>
				<td align="right">0</td>
		<%	END IF
			END IF 
			IF ichkNull = 0 THEN
			For intC = intI To  (intB-1) '사업부서 수만큼 루프
					IF arrBiz(0,intC)  = arrList(2,intLoop) and arrBiz(2,intC) = arrList(11,intLoop) THEN '사업부서와 손익디비 부서가 동일할때 값 보여준다.  
					 	sumM = sumM + arrList(4,intLoop)-arrList(3,intLoop) 
					  arrgrp(intC+1) = arrgrp(intC+1) + arrList(4,intLoop)-arrList(3,intLoop) 
					 	arrup(intC+1) = arrup(intC+1)+ arrList(4,intLoop)-arrList(3,intLoop) 
				%>
				<td align="right"><%IF arrList(11,intLoop) THEN%><a href="javascript:showProfitBizDetail('<%=arrList(2,intLoop)%>','<%=arrList(9,intLoop)%>');"><%END IF%><%=formatnumber(arrList(4,intLoop)-arrList(3,intLoop),0)%></a></td>
				<%	 intI = intC+1
						Exit For  
					ELSE	'손익디비에 값 없을떄 0 뿌려준다.
				%>
				<td align="right">0</td>
				<%		
					END IF
			Next%> 
		<%END IF	
				oldacc = accType 
				oldbizType = arrList(5,intLoop)
				oldgrp = arrList(1,intLoop)
				 
				if arrList(6,intLoop) = "" THEN '계정분류 없는 경우 표시 안한다
						oldup = ""
				else	
					oldup = arrList(7,intLoop)	
				end if
			Next  
			For intC= intI To (intB-1)
		%>
				<td align="right">0</td>
		<%Next %>  
				<td  align="right"><%=formatnumber(sumM,0)%></td>
			</tr>
			<%  
			 
						if  oldup <> "" THEN
							sumup = 0
				%>
				<tr bgcolor="#EEEEEE"> 
				<td colspan="5" align="center"><%=oldup%></td> 
					<%For intSum = 0 To intB
						sumup = sumup + arrup(intSum)
				%>
				<td  align="right"><%=formatnumber(arrup(intSum),0)%></td> 
				<%Next%>
					<td  align="right"><%=formatnumber(sumup,0)%></td> 
				</tr>
				<%	
				END IF
				 
				%>	
				<%  
						sumgrp = 0
				%>
				<tr bgcolor="#DDFFDD"> 
				<td colspan="5" align="center" style="border-bottom:2px solid <%=adminColor("tablebg")%>;"><%=oldgrp%></td> 
				<%For intSum = 0 To intB
						sumgrp = sumgrp + arrgrp(intSum)
				%>
				<td  align="right"  style="border-bottom:2px solid <%=adminColor("tablebg")%>;"><%=formatnumber(arrgrp(intSum),0)%></td> 
				<%Next%>
					<td  align="right"  style="border-bottom:2px solid <%=adminColor("tablebg")%>;"><%=formatnumber(sumgrp,0)%></td> 
				</tr>  
			 
		<%ELSE%>
		<tR>
			<td bgcolor="#FFFFFF" colspan="35" align="center">등록된 내용이 없습니다.</td>
		</tr>	
		<%END IF%>
		</table>
	</td>
</tr>
</table>	
<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Page : /admin/eventmanage/event_schedule.asp
' Description :  이벤트 스케쥴관리
' History : 2007.02.22 정윤정 생성
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/eventmanage/common/event_function.asp"-->
<!-- #include virtual="/lib/classes/event/eventmanageCls.asp"-->
<%
'변수선언
Dim intY, intM, strY,strM, intD
Dim FirstDate, LastDate, LastDay
Dim iTotCnt, arrList,intLoop
Dim iPageSize, iCurrpage ,iDelCnt
Dim iStartPage, iEndPage, iTotalPage, ix,iPerCnt	
dim cEvtList
Dim sCategory 
Dim arreventstate,sState
		
'파라미터값 받기 & 기본 변수 값 세팅
	iCurrpage = Request("iC")	'현재 페이지 번호
	IF iCurrpage = "" THEN
		iCurrpage = 1	
	END IF	  
	iPageSize = 20		'한 페이지의 보여지는 열의 수
	iPerCnt = 10		'보여지는 페이지 간격
	
	sCategory = Request("selC")	 
	sState = Request("selState")	
 	IF sState = "" THEN sState = -1	
	strY = Request("selY")
	strM = Request("selM")
	If strY = "" THEN
		strY = year(date())
	END IF	
	If strM = "" THEN
		strM = month(date())
	END IF	

	' 현재 달의 월말 값 계산
	IF strM = 4 OR strM =6 OR strM = 9 OR strM = 11 THEN		
		LastDay = 30
	ELSEIF strM = 2 AND NOT (strY MOD 4) = 0 THEN
		LastDay = 28
	ELSEIF strM = 2 AND (strY MOD 4) = 0 THEN
		IF (strY MOD 100) = 0 THEN
			IF (strY MOD 400) = 0 THEN
				LastDay = 29
			ELSE
				LastDay = 30
			END IF
		ELSE
			LastDay = 29
		END IF
	ELSE
		LastDay = 31
	END IF

	'시작일
	FirstDate = DateSerial(strY,strM,1)
	'종료일
	LastDate = DateSerial(strY,strM,LastDay)

	'데이터 가져오기
	 set cEvtList = new ClsEventSchedule
	 	 cEvtList.FCPage = iCurrpage	'현재페이지
		 cEvtList.FPSize = iPageSize '한페이지에 보이는 레코드갯수 
		 
		 cEvtList.FSCategory= sCategory
		 cEvtList.FSState	= sState		
	 	 cEvtList.FFDate 	= FirstDate
	 	 cEvtList.FLDate  	= LastDate
	 	 
	 	 
		 arrList = cEvtList.fnGetList
		 iTotCnt = cEvtList.FTotCnt	'전체 데이터  수
	 set cEvtList = nothing
	 
	 arreventstate = fnSetCommonCodeArr("eventstate",False) 
	 
 	iTotalPage 	=  Int(iTotCnt/iPageSize)	'전체 페이지 수
	IF (iTotCnt MOD iPageSize) > 0 THEN	iTotalPage = iTotalPage + 1		
%>
<script language="javascript">
<!--
	function jsGoPage(iP){
		document.frmEvt.iC.value = iP;
		document.frmEvt.submit();	
	}	
	
	function jsChEvt(){	
		if(document.frmEvt.chkE.checked){
		  document.frmEvt.iE.value = 1; 
		}else{
		  document.frmEvt.iE.value = 0; 
		}
		document.frmEvt.submit();	
	}
	
	function jsViewDetail(eC,sMod){
	  if (sMod == 0){
		eval("document.all.sD"+eC).style.display = "";
	  }else{
	  	eval("document.all.sD"+eC).style.display = "none";
	  }	
	}
//-->
</script>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frmEvt" method="post"> 
	<input type="hidden" name="iC">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
		<td align="left">
			<select class="select" name="selY" onChange="document.frmEvt.submit();">
				<%For intY = year(date())+1 To 2006 Step -1 %>
				<option value="<%=intY%>" <%IF Cint(strY) = Cint(intY) THEN%>selected<%End IF%>><%=intY%></option>
				<%Next%>
			</select>년
			<select class="select" name="selM" onChange="document.frmEvt.submit();">
				<%For intM = 1 To 12  %>
				<option value="<%=intM%>" <%IF Cint(strM) = Cint(intM) THEN%>selected<%End IF%>><%=intM%></option>
				<%Next%>
			</select>월
			&nbsp;
			카테고리:
			<select name="selC" onChange="document.frmEvt.submit();">
			<% sbGetOnlyOptCategoryLarge sCategory%>
			<option value="-1" <%if sCategory = "-1" then%>selected<%end if%>>기타</option>
			</select>
			&nbsp;		
			상태:
			<select name="selState" onChange="document.frmEvt.submit();">
			<option value="-1" <%IF sState = "-1" THEN%>selected<%END IF%>>종료이벤트 제외</option>
			<%IF isArray(arreventstate) THEN
				For intLoop = 0 To UBound(arreventstate,2)
				%>
			<option value="<%=arreventstate(0,intLoop)%>" <%If CStr(sState) = CStr(arreventstate(0,intLoop)) THEN%>selected<%END IF%>><%=arreventstate(1,intLoop)%></option>
			<%	Next
			END IF%>
			</select>				
		</td>
		
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
		</td>
	</tr>
	</form>
</table>
<!-- 검색 끝 -->

<p>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="40">
		검색결과 : <b><%= iTotCnt %></b>
		&nbsp;
		페이지 : <b><%= iCurrpage %> / <%= iTotalPage %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="40">코드</td>
	<td width="40">관련<br>코드</td>
	<td width="60">상태</td>
	<td>카테고리</td>
	<td>이벤트명</td>
	<%For intD = 1 To LastDay%>
	<td width="12" align="center"><%=intD%></td>
	<%Next%>
</tr>	
<% dim tsday, teday, tmpday, intLD
   dim strBg, ThisDate
IF isArray(arrList) THEN
	For intLoop = 0 To UBound(arrList,2) 	
	  tsday = Replace(arrList(7,intLoop),"-","")
	  teday = Replace(arrList(8,intLoop),"-","")
%>
<tr bgcolor="FEFEFE" height="10" onClick="javascript:opener.location.href='/admin/eventmanage/event/V2/event_modify.asp?eC=<%=arrList(0,intLoop)%>'" style="cursor:hand;">
	<td align="center"><%=arrList(0,intLoop)%></td>
	<td align="center"><font color="#666666"><%=arrList(10,intLoop)%></font></td>
	<td align="center"><%=fnSetStatusDesc(arrList(6,intLoop),arrList(7,intLoop),arrList(8,intLoop), fnGetCommCodeArrDesc(arreventstate, arrList(6,intLoop)))%></td>
	<td align="center"><%=arrList(9,intLoop)%></td>
	<td><%=db2html(arrList(4,intLoop))%></td>
	<% For intLD = 1 To LastDay 
	    tmpday = DateSerial(strY,strM,intLD)
	    tmpday =Replace(tmpday,"-","")
	    ThisDate = Replace(date(),"-","")
	    
	 if (tsday <= tmpday and teday >= tmpday)then 
	 	if (tmpday<ThisDate) then
	 		strBg = " background=""/images/dot4.gif"""
	 	else	
			strBg = " background=""/images/dot40.gif"""
		end if	
	 else 
	  	strBg = ""
	 end if%>	
	<td <%=strBg%>></td>	
<%Next%>
</tr>	
<%  Next
END IF%>


<!-- 페이징처리 -->
<%		
iStartPage = (Int((iCurrpage-1)/iPerCnt)*iPerCnt) + 1	

If (iCurrpage mod iPerCnt) = 0 Then																
	iEndPage = iCurrpage
Else								
	iEndPage = iStartPage + (iPerCnt-1)
End If	
%>


	<tr height="25" bgcolor="FFFFFF">
		<td colspan="40" align="center">
			<% if (iStartPage-1 )> 0 then %><a href="javascript:jsGoPage(<%= iStartPage-1 %>)" onfocus="this.blur();">[pre]</a>
			<% else %>[pre]<% end if %>
	        <%
				for ix = iStartPage  to iEndPage
					if (ix > iTotalPage) then Exit for
					if Cint(ix) = Cint(iCurrpage) then
			%>
				<a href="javascript:jsGoPage(<%= ix %>)" class="menu_link3" onfocus="this.blur();"><font color="00abdf"><strong><%=ix%></strong></font></a>
			<%		else %>
				<a href="javascript:jsGoPage(<%= ix %>)" class="menu_link3" onfocus="this.blur();"><%=ix%></a>
			<%
					end if
				next
			%>
	    	<% if Cint(iTotalPage) > Cint(iEndPage)  then %><a href="javascript:jsGoPage(<%= ix %>)" onfocus="this.blur();">[next]</a>
			<% else %>[next]<% end if %>
		</td>
	</tr>
    
    </form>
</table>
<!-- #include virtual="/lib/db/dbclose.asp" -->
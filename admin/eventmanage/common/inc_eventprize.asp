<%
'###########################################################
' Page : /admin/eventmanage/common/eventprize_regist.asp
' Description : 당첨자 등록처리 include
' History : 2007.02.13 정윤정 생성
'###########################################################

 Dim cEvtPrize
 Dim arrPrize, intLoop
 Dim iTotCnt
 Dim iPageSize, iCurrpage ,iDelCnt
 Dim iStartPage, iEndPage, iTotalPage, ix,iPerCnt
 Dim arrPrizeType, arrPrizeStatus
  
 IF egKindCode = "" THEN egKindCode = 0	
	
 iCurrpage 	= requestCheckVar(Request("iC"),10)		 '현재 페이지 번호
 IF iCurrpage = "" THEN	iCurrpage = 1
	
	iPageSize = 30		'한 페이지의 보여지는 열의 수
	iPerCnt = 10		'보여지는 페이지 간격
	
	set cEvtPrize = new ClsEventPrize
	cEvtPrize.FECode	  	= eCode			'이벤트 코드
	cEvtPrize.FEGKindCode 	= egKindCode	'그룹코드(핑거스,문화이벤트 회차)
	cEvtPrize.FCPage 		= iCurrpage
	cEvtPrize.FPSize 		= iPageSize
	arrPrize = cEvtPrize.fnGetPrize		'당첨내역
	iTotCnt = cEvtPrize.FTotCnt			'전체 데이터  수
	set cEvtPrize = nothing
	arrPrizeType = fnSetCommonCodeArr("evtprizetype",False)
	arrPrizeStatus= fnSetCommonCodeArr("evtprizestatus",False)
	
	iTotalPage 	=  int((iTotCnt-1)/iPageSize) +1  '전체 페이지 수	
%>
<script type="text/javascript" src="/js/jquery-1.7.2.min.js"></script>
<script language="javascript">
<!--
//당첨자 등록
  function jsSetWinner(eC,egKC,epC){
  	var winW, popURL;
  	if (epC > 0){
  		popURL ="/admin/eventmanage/event/pop_event_changewinner.asp?epC="+epC;  		
  	}else{
  		popURL="/admin/eventmanage/event/pop_event_winner.asp?eC="+eC+"&egKC="+egKC;
  	}
  	winW = window.open(popURL,'popW','width=1000, height=700, scrollbars=yes');
  	winW.focus();
  }
  
  //당첨 취소
  
  	//페이징처리
		function jsGoPage(iP){
		document.frmPrize.iC.value = iP;
		document.frmPrize.submit();
	}

function tnCheckAll(bool, comp){
    var frm = comp.form;
    if (!comp.length){
        comp.checked = bool;
    }else{
        for (var i=0;i<comp.length;i++){
            comp[i].checked = bool;
        }
    }
}

function jsSMSSendPop(){
	if($("input:checkbox[name='cksel']:checked").length<1){
		alert("발송대상을 선택해주세요.");
	}else{
		frm = document.frmPrize;
		window.open('', 'popSMS', 'width=500, height=700');
		frm.action = "/admin/eventmanage/common/pop_prize_sms_send.asp";
		frm.target = "popSMS";
		frm.method = "post";
		frm.submit();
	}
}
//-->
</script>
<table width="100%" border="0" align="left" class="a" cellpadding="0" cellspacing="1">
	<tr>
		<td>
			<table width="100%" border="0" align="left" class="a" cellpadding="3" cellspacing="1">
			<tr>
				<td>
				<input type="button" name="btnadd"  value="SMS 전송" onClick="javascript:jsSMSSendPop();" class="button">
				<input type="button" name="btnadd"  value="새 당첨등록" onClick="javascript:jsSetWinner(<%=eCode%>,<%=egKindCode%>,0);" class="button">
				</td>
			</tr>	
			</table>
		</td>	
	<tr>
		<td>
			<table width="100%" border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
			<form name="frmPrize" method="post" >
			<input type="hidden" name="menupos" value="<%=menupos%>">
			<input type="hidden" name="iC" value="<%=iCurrpage%>">
			<input type="hidden" name="eC" value="<%=eCode%>">	
			<input type="hidden" name="egKC" value="<%=egKindCode%>">			
			<tr bgcolor="#FFFFFF" height="25">
				<td colspan="10">검색결과 : <b><%=iTotCnt%></b>&nbsp;&nbsp;페이지 : <b><%=iCurrpage%> / <%=iTotalPage%></b></td>
			</tr>		
			<tr>
				<td align="center"  width="50" bgcolor="<%= adminColor("tabletop") %>"><input type="checkbox" onClick="tnCheckAll(this.checked,frmPrize.cksel);" /></td>
				<td align="center"  width="50" bgcolor="<%= adminColor("tabletop") %>">당첨코드</td>							
				<td align="center"  width="50" bgcolor="<%= adminColor("tabletop") %>">등수</td>
				<td align="center"  bgcolor="<%= adminColor("tabletop") %>">등수별칭</td>
				<td align="center"  width="70" bgcolor="<%= adminColor("tabletop") %>">구분</td>
				<td align="center"  bgcolor="<%= adminColor("tabletop") %>">사은품명(상품번호)</td>							
				<td align="center"  width="100" bgcolor="<%= adminColor("tabletop") %>">당첨자</td>
				<td align="center"  width="150"  bgcolor="<%= adminColor("tabletop") %>">당첨확인기간</td>
				<td align="center"  width="100" bgcolor="<%= adminColor("tabletop") %>">상태</td>				
				<td align="center"  width="60" bgcolor="<%= adminColor("tabletop") %>">양도<br>당첨코드</td>
			</tr>
			<%IF isArray(arrPrize) THEN%>	
				<%For intLoop = 0 To UBound(arrPrize,2)	%>
				<tr>
					<td bgcolor="#FFFFFF" align="center"><input type='checkbox' name='cksel' id="cksel<%=intLoop%>" value='<%=arrPrize(0,intLoop)%>' /></td>
					<td bgcolor="#FFFFFF" align="center"><%=arrPrize(0,intLoop)%></td>
					<td bgcolor="#FFFFFF" align="center"><%=arrPrize(1,intLoop)%></td>
					<td bgcolor="#FFFFFF" align="center"><%=arrPrize(2,intLoop)%></td>
					<td bgcolor="#FFFFFF" align="center"><%=fnGetCommCodeArrDesc(arrPrizeType,arrPrize(14,intLoop))%></td>
					<td bgcolor="#FFFFFF"  align="left">&nbsp;<%=arrPrize(11,intLoop)%><%IF arrPrize(13,intLoop) <> 0 THEN%>(<%=arrPrize(13,intLoop)%>)<%END IF%></td>
					<td bgcolor="#FFFFFF"  align="center"><%=arrPrize(5,intLoop)%></td>
					<td bgcolor="#FFFFFF" align="left">&nbsp;<%if arrPrize(7,intLoop) <> "1900-01-01" then%><%=arrPrize(7,intLoop)%> ~ <%if arrPrize(8,intLoop) <> "1900-01-01" then%><%=arrPrize(8,intLoop)%><%end if%><%end if%></td>
					<td bgcolor="#FFFFFF" align="center">
						<%IF arrPrize(9,intLoop) = 5 THEN %>							
							<input type="button" class="button" value="양도신청" onClick="jsSetWinner(<%=eCode%>,<%=egKindCode%>,<%=arrPrize(0,intLoop)%>);">
						<%ELSEIF datediff("d",arrPrize(8,intLoop),date()) > 0 AND  arrPrize(9,intLoop) = 0 THEN%>								
							<input type="button" class="button" value="기간만료" onClick="jsSetWinner(<%=eCode%>,<%=egKindCode%>,<%=arrPrize(0,intLoop)%>);">
						<%ELSE%>	
							<%=fnGetCommCodeArrDesc(arrPrizeStatus,arrPrize(9,intLoop))%>
						<%END IF%>		
						</td>	
					<td bgcolor="#FFFFFF" align="center"><%=arrPrize(15,intLoop)%></td>
				</tr>	
				<%Next%>				
			<%else%>	
				<tr>
					<td bgcolor="#FFFFFF" colspan="10" align="center">당첨내역이 없습니다.</td>
				</tr>
			<%END IF%>	
			</table>	
		</td>
			
	</tr>
		
		
	<tr>
		<td>
			<!-- 페이징처리 -->
			<%
			iStartPage = (Int((iCurrpage-1)/iPerCnt)*iPerCnt) + 1
			
			If (iCurrpage mod iPerCnt) = 0 Then
				iEndPage = iCurrpage
			Else
				iEndPage = iStartPage + (iPerCnt-1)
			End If
			%>
			<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
			    <tr valign="bottom" height="25">        
			        <td valign="bottom" align="center">
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
		</td>
	</tr>
</table>	
<%@ language=vbscript %>
<% option explicit %>
<% response.Charset="euc-kr" %> 
<%
'###########################################################
' Description : 보낸결재함 리스트
' History : 2011.03.14 정윤정 생성
'			2019.05.27 한용민 수정
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" --> 
<!-- #include virtual="/lib/db/dbopen.asp" --> 
<!-- #include virtual="/lib/util/htmllib.asp"-->  
<!-- #include virtual="/lib/classes/approval/eappCls.asp"-->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenMemberCls.asp"-->
<% 
Dim clseapp
Dim sadminId,ireportState ,ireportidx
Dim iCurrpage,ipagesize,iTotCnt,iTotalPage
Dim arrList,intLoop , reportname, reportcontents, reportprice, regdate, username
	reportname =  requestCheckvar(Request("reportname"),120)
	reportcontents =  requestCheckvar(Request("reportcontents"),120)
	reportprice =  requestCheckvar(getNumeric(Request("reportprice")),10)
	regdate =  requestCheckvar(Request("regdate"),10)
	username =  requestCheckvar(Request("username"),32)
	iPageSize = 30 
	iCurrPage = requestCheckvar(Request("iCP"),10)
	if iCurrPage="" then iCurrPage=1 
	 
	sadminId =  session("ssBctId")
 	ireportState =  requestCheckvar(Request("iRS"),4) 
 	ireportidx =  requestCheckvar(Request("iridx"),10)
 	IF ireportidx = "" THEN ireportidx = 0
 		
'결재 기본 폼 정보 가져오기
set clseapp = new CEApproval
	clseapp.FadminId 	= sadminId
	clseapp.FreportState= ireportState
	clseapp.FCurrpage 	= iCurrpage
	clseapp.FPagesize	= ipagesize
	clseapp.freportname	= reportname
	clseapp.freportcontents	= reportcontents
	clseapp.freportprice	= reportprice
	clseapp.fregdate	= regdate
	clseapp.fusername	= username
	arrList = clseapp.fnGetEAppSendList
	iTotCnt = clseapp.FTotCnt 
set clseapp = nothing
	iTotalPage 	=  int((iTotCnt-1)/iPageSize) +1  '전체 페이지 수
	
 Dim iRectMenu 
 iRectMenu = requestCheckvar(Request("iRM"),10) 
  
%>
<html>
<head> 
<!-- #include virtual="/admin/approval/eapp/eappheader.asp"-->  
<script type="text/javascript" src="/js/ajax.js"></script> 
<script type="text/javascript" src="eapp.js"></script>  
<script type="text/javascript" src="/js/jsCal/js/jscal2.js"></script>
<script type="text/javascript" src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script language="javascript">
	function jsMod(iridx){
		top.eappDetail.location.href = "modeapp.asp?iridx="+iridx+"&iRM=<%=iRectMenu%>"; 
	} 
	
	 initializeReturnFunction("processAjax()");
   initializeErrorFunction("onErrorAjax()"); 
   
   var _reportidx = 0;
    function processAjax(){
        var reTxt = xmlHttp.responseText;  
       	eval("document.all.dPR"+_reportidx).innerHTML = reTxt; 
    }
    
    function onErrorAjax() {
       alert("ERROR : " + xmlHttp.status);
    }
      
   function jsViewPayRequest(iridx){     
   		_reportidx = iridx;
    	initializeURL('ajaxPayRequest.asp?iridx='+iridx);
    	startRequest();  
   }
   
   function jsReSetHtm(iridx){ 
   	eval("document.all.dPR"+iridx).innerHTML = "";  
   }
	function frmsubmit(page){
		frmList.iCP.value = page;
		frmList.submit();
	}
</script>
</head>
<body leftmargin="0" topmargin="0">
<div style="height:100%;overflow-y:auto;">
<table width="100%" height="100%" cellpadding="0" cellspacing="0"  border="0">
<tr> 
	<td valign="top" >
		<table width="100%" cellpadding="0" cellspacing="1" class="a" border="0"> 
		<form name="frmList" method="get">
		<input type="hidden" name="iCP">
		<input type="hidden" name="iRS" value="<%=ireportstate%>">
		<input type="hidden" name="iridx" value="<%=ireportidx%>">
		<tr> 
			<td height="25"><font color="#4E9FC6"><b>보낸결재함 >결재문서 <%IF ireportState >= 0 THEN%>><%=fnGetMenu("S1",ireportState,"")%><%END IF%></b></font></td>
		</tr> 
		<tr>
			<td>
				<!-- 검색 시작 -->
				<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
				<tr align="center" bgcolor="#FFFFFF" >
					<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
					<td align="left">
						* 품의서명 : <input type="text" name="reportname" id="reportname" value="<%= reportname %>" size="15" maxlength=120 class="text">
						* 내용 : <input type="text" name="reportcontents" id="reportcontents" value="<%= reportcontents %>" size="15" maxlength=120 class="text">
						<Br>
						* 픔의금액 : <input type="text" name="reportprice" id="reportprice" value="<%= reportprice %>" size="8" maxlength=10 class="text">
						* 작성자 : <input type="text" name="username" id="username" value="<%= username %>" size="8" maxlength=10 class="text">
					</td>
					<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
						<input type="button" class="button_s" value="검색" onClick="frmsubmit('1');">
					</td>
				</tr>
				<tr align="center" bgcolor="#FFFFFF" >
					<td align="left">
						* 작성일 : <input type="text" id="termSdt" name="regdate" size="8" maxlength=10 value="<%= regdate %>" />
						<img src="/images/admin_calendar.png" alt="달력으로 검색" id="ChkStart_trigger" onclick="return false;" />
						<script type="text/javascript">
							var CAL_Start = new Calendar({
								inputField : "termSdt", trigger    : "ChkStart_trigger",
								onSelect: function() {
									var date = Calendar.intToDate(this.selection.get());
									//CAL_End.args.min = date;
									//CAL_End.redraw();
									this.hide();
								}, bottomBar: true, dateFormat: "%Y-%m-%d" <%'=chkIIF(regdate<>"",", max: " & replace(regdate,"-",""),"")%>
							});
						</script>
					</td>
				</tr>
				</table>
				<!-- 검색 끝 -->
				<Br>
			</td>
	 	</tr>
		<tr>
			<td>  
				<table width="100%" align="left" cellpadding="5" cellspacing="1" class="a" border="0" bgcolor="#cccccc">
				<tr bgcolor="<%= adminColor("tabletop") %>" align="center">
					<td width="20">Idx</td>  
					<td>품의서명</td>
					<td nowrap>품의금액</td>  
					<td>작성일</td>  
					<td nowrap>상태</td>  
					<%IF ireportState >= 7 THEN%> 
					<td nowrap>결제<br>요청서</td>
					<%END IF%>
				</tr>
				<%
					IF isArray(arrList) THEN
						For intLoop = 0 To UBound(arrList,2)
				%>
				<tr id="t<%=arrList(0,intLoop)%>" bgcolor="#FFFFFF" onclick="jsMod(<%=arrList(0,intLoop)%>);ChangeColor(this,'#CEF6EC','FFFFFF');" style="cursor:hand;"> 
					<td align="center"><%=arrList(0,intLoop)%></td>  
					<td><%=arrList(1,intLoop)%></td>
					<td align="right" nowrap><%=formatnumber(arrList(2,intLoop),0)%></td> 
					<td nowrap><%=FormatDate(arrList(10,intLoop),"0000-00-00")%></td>
					<td align="center" nowrap><%=fnGetReportState(arrList(11,intLoop))%></a></td>
					<%IF ireportState >= 7 THEN %>
					<td align="center" nowrap>
						 <% IF arrList(8,intLoop) THEN %>
						 <%IF arrList(9,intLoop) < arrList(2,intLoop) THEN%><font color="blue">등록</font><%ELSE%><a href="javascript:jsViewPayRequest(<%=arrList(0,intLoop)%>);">등록<Br>완료</a><%END IF%> 
						 <div style="position:relative;background-color:#eeeeee"> 
						 <div id="dPR<%=arrList(0,intLoop)%>" style="position:absolute;left:-240px;top:-10px;z-index:100;background-color:white;"></div>
						 </div>
						 <%END IF%>
					</td>
					<%END IF%>
				</tr> 
				<%	
					Next
					ELSE	
				%>
				<tr bgcolor="#FFFFFF">
					<td colspan="5" align="center">등록된 내역이 없습니다.</td>
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
				<td>
					<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" >
					    <tr height="25">        
					        <td  align="center">
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
		</form>
			</table>   
	</td> 
</tr>	 
</table>  
</div>
</body>
</html>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 결제요청서 결재선 리스트
' History : 2011.03.14 정윤정 생성
'			2019.05.27 한용민 수정
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" --> 
<!-- #include virtual="/lib/db/dbopen.asp" --> 
<!-- #include virtual="/lib/util/htmllib.asp"--> 
<!-- #include virtual="/lib/classes/approval/payrequestCls.asp"-->
<!-- #include virtual="/lib/classes/approval/eappCls.asp"-->
<%
Dim clsPay
Dim sadminId,ipayrequeststate,iauthstate,blnLast
Dim ireportidx, ipayrequestidx
Dim iCurrpage,ipagesize,iTotCnt,iTotalPage
Dim arrList,intLoop , reportname, payrequestprice, paydate, username
	reportname =  requestCheckvar(Request("reportname"),120)
	payrequestprice =  requestCheckvar(getNumeric(Request("payrequestprice")),10)
	paydate =  requestCheckvar(Request("paydate"),10)
	username =  requestCheckvar(Request("username"),32)
	iPageSize = 20
	iCurrPage = requestCheckvar(Request("iCP"),10)
	if iCurrPage="" then iCurrPage=1 
	 
	sadminId =  session("ssBctId") 
  
'결재 기본 폼 정보 가져오기
set clsPay = new CPayRequest  
	clsPay.FadminID		= sadminId
	clsPay.FCurrpage 	= iCurrpage
	clsPay.FPagesize	= ipagesize
	clsPay.freportname	= reportname
	clsPay.fpayrequestprice	= payrequestprice
	clsPay.fpaydate	= paydate
	clsPay.fusername	= username
	arrList = clsPay.fnGetPayRequestAuthLine
	iTotCnt = clsPay.FTotCnt 
set clsPay = nothing
	iTotalPage 	=  int((iTotCnt-1)/iPageSize) +1  '전체 페이지 수
	
	Dim iRectMenu
	iRectMenu="M120"
%>
<html>
<head> 
<!-- #include virtual="/admin/approval/eapp/eappheader.asp"-->  
<script type="text/javascript" src="/admin/approval/eapp/eapp.js"></script>  
<script type="text/javascript" src="/js/jsCal/js/jscal2.js"></script>
<script type="text/javascript" src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script language="javascript">
<!--  
function jsConfirm(reportidx,payrequestidx){
		top.eappDetail.location.href = "viewpayrequest.asp?iridx="+reportidx+"&ipridx="+payrequestidx;
	}
	function frmsubmit(page){
		frmList.iCP.value = page;
		frmList.submit();
	}
//-->
</script> 
</head>
<body leftmargin="0" topmargin="0">
<div style="height:100%;overflow-y:auto;">
<table width="100%" height="100%" cellpadding="0" cellspacing="0"  border="0">
<tr> 
	<td valign="top">
		<form name="frmList" method="post" action="">
		<input type="hidden" name="iCP" value="<%=iCurrPage%>">   
		<input type="hidden" name="iridx" value="<%=ireportidx%>">
		<input type="hidden" name="ipridx" value="<%=ipayrequestidx%>">
		<table width="100%" cellpadding="0" cellspacing="1" class="a" border="0"> 
		<tr> 
			<td height="25"><font color="#4E9FC6"><b>결제요청서> <%=fnGetMenu("R2","","")%></td>
		</tr>
		<tr>
			<td>
				<!-- 검색 시작 -->
				<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
				<tr align="center" bgcolor="#FFFFFF" >
					<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
					<td align="left">
						* 결제요청서명 : <input type="text" name="reportname" id="reportname" value="<%= reportname %>" size="15" maxlength=120 class="text">
						<Br>
						* 결제요청금액 : <input type="text" name="payrequestprice" id="payrequestprice" value="<%= payrequestprice %>" size="8" maxlength=10 class="text">
						* 작성자 : <input type="text" name="username" id="username" value="<%= username %>" size="8" maxlength=10 class="text">
					</td>
					<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
						<input type="button" class="button_s" value="검색" onClick="frmsubmit('1');">
					</td>
				</tr>
				<tr align="center" bgcolor="#FFFFFF" >
					<td align="left">
						* 결제일 : <input type="text" id="termSdt" name="paydate" size="8" maxlength=10 value="<%= paydate %>" />
						<img src="/images/admin_calendar.png" alt="달력으로 검색" id="ChkStart_trigger" onclick="return false;" />
						<script type="text/javascript">
							var CAL_Start = new Calendar({
								inputField : "termSdt", trigger    : "ChkStart_trigger",
								onSelect: function() {
									var date = Calendar.intToDate(this.selection.get());
									//CAL_End.args.min = date;
									//CAL_End.redraw();
									this.hide();
								}, bottomBar: true, dateFormat: "%Y-%m-%d" <%'=chkIIF(paydate<>"",", max: " & replace(paydate,"-",""),"")%>
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
						<table width="100%" align="left" cellpadding="5" cellspacing="1" class="a"   border="0" bgcolor="#cccccc">
							<tr bgcolor="<%= adminColor("tabletop") %>" align="center">
							<td rowspan="2" width="50">결제요청<br>Idx</td> 
							<td rowspan="2">결제요청서명</td> 
							<td  rowspan="2" nowrap>결제요청금액</td>
							<td nowrap width="60">결제요청일</td>
							<td nowrap width="60">결제일</td>
							<td nowrap rowspan="2">작성자</td>
						</tr>
						<tr bgcolor="<%= adminColor("tabletop") %>" align="center">	
							<td nowrap>결제예정일</td>
							<td nowrap>서류제출</td>   
						</tr> 
						<%IF isArray(arrList) THEN
							For intLoop = 0 To UBound(arrList,2)
						%>
						<tr  id="t<%=arrList(0,intLoop)%>" bgcolor="#FFFFFF" align="center" onclick="jsConfirm(<%=arrList(8,intLoop)%>,<%=arrList(0,intLoop)%>);evalChangeColor('document.all.t<%=arrList(0,intLoop)%>','#CEF6EC','FFFFFF');" style="cursor:hand;"> 
							<td rowspan="2"><%=arrList(0,intLoop)%></td>
							<td rowspan="2">[<%=arrList(8,intLoop)%>] <%=arrList(6,intLoop)%></td> 
							<td rowspan="2"><%=formatnumber(arrList(2,intLoop),0)%></td>
							<td><%IF arrList(1,intLoop) <> "" THEN%><%=formatdate(arrList(1,intLoop),"0000-00-00")%><%END IF%></td>  
							<td><%IF arrList(3,intLoop) <> "" THEN%><%=formatdate(arrList(3	,intLoop),"0000-00-00")%><%END IF%></td> 
							<td rowspan="2"><%=arrList(10,intLoop)%></td>
						</tr>
						<tr id="t<%=arrList(0,intLoop)%>" bgcolor="#FFFFFF" align="center" onclick="jsConfirm(<%=arrList(8,intLoop)%>,<%=arrList(0,intLoop)%>);evalChangeColor('document.all.t<%=arrList(0,intLoop)%>','#CEF6EC','FFFFFF');" style="cursor:hand;">	 
							<td nowrap><%IF arrList(4,intLoop) <> "" THEN%><%=formatdate(arrList(4,intLoop),"0000-00-00")%><%END IF%></td>
							<td nowrap><%IF arrList(11,intLoop) THEN%>Y<%ELSE%>N<%END IF%></td>   
						</tr> 
						<%	
							Next
							ELSE	
						%>
						<tr bgcolor="#FFFFFF">
							<td colspan="8" align="center">등록된 내역이 없습니다.</td>
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
							<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
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
			<!----------------- /내용 ---------------------------> 
			</td>
		</tr>	 
		</table>
		</form>
	</td>
</tr>
</table>
</div>
<!-- 페이지 끝 -->
</body>
</html>
<!-- #include virtual="/lib/db/dbclose.asp" -->
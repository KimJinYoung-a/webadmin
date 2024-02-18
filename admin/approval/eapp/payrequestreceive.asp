<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 결제요청서 리스트
' History : 2011.03.14 정윤정  생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/approval/payrequestCls.asp"-->
<!-- #include virtual="/lib/classes/approval/eappCls.asp"-->
<!-- #include virtual="/lib/classes/approval/payManagerCls.asp"-->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenDepartmentCls.asp"-->
<%
Dim clsPay
Dim sadminId,ipayrequeststate,iauthstate,blnLast
Dim ireportidx, ipayrequestidx
Dim iCurrpage,ipagesize,iTotCnt,iTotalPage
Dim reportname, reportprice, regdate, username, department_id
	reportname =  requestCheckvar(Request("reportname"),300)
	reportprice	= requestCheckvar(getNumeric(Request("reportprice")),10)
	regdate =  requestCheckvar(Request("regdate"),10)
	username =  requestCheckvar(Request("username"),32)
	department_id = requestCheckvar(Request("department_id"),10)

Dim arrList,intLoop
Dim iPayRequeststate000 ,iPayRequeststate001 ,iPayRequeststate110 ,iPayRequeststate111 ,iPayRequeststate710 ,iPayRequeststate711
Dim iPayRequeststate970 ,iPayRequeststate971 ,iPayRequeststate550 ,iPayRequeststate551
Dim clsPM, arrPM, intP  ,blnMod, igbn,iRectMenu
	iPageSize = 20
	iCurrPage = requestCheckvar(Request("iCP"),10)
	if iCurrPage="" then iCurrPage=1

	 blnMod = 0
	 igbn			= requestCheckvar(Request("igbn"),1)
	sadminId =  session("ssBctId")
 	ipayrequeststate= requestCheckvar(Request("iPRS"),4)
 	iauthstate		= requestCheckvar(Request("iAS"),4)
 	blnLast			= requestCheckvar(Request("blnL"),1)
 iRectMenu =	requestCheckvar(Request("iRM"),10)
 	ireportidx 		=  requestCheckvar(Request("iridx"),10)
	ipayrequestIdx	= requestCheckvar(Request("ipridx"),10)
	if ireportidx = "" THEN ireportidx = 0
		if ipayrequestIdx = "" THEN ipayrequestIdx = 0
	if ipayrequeststate = "" THEN ipayrequeststate = 1
	if iauthstate = "" THEN iauthstate = 0
'결재 기본 폼 정보 가져오기
set clsPay = new CPayRequest
	clsPay.Fpayrequeststate = ipayrequeststate
	clsPay.Fauthstate	= iauthstate
	clsPay.FisLast		= blnLast
	clsPay.FCurrpage 	= iCurrpage
	clsPay.FPagesize	= ipagesize
	clsPay.Freportname	= reportname
	clsPay.Freportprice	= reportprice
	clsPay.Fregdate	= regdate
	clsPay.Fusername	= username
	clsPay.Fdepartment_id = department_id
	arrList = clsPay.fnGetPayRequestReceiveList
	iTotCnt = clsPay.FTotCnt
set clsPay = nothing
	iTotalPage 	=  int((iTotCnt-1)/iPageSize) +1  '전체 페이지 수

Set clsPM	= new CPayManager
	arrPM	= clsPM.fnGetPayManager
Set clsPM 	= nothing

	IF isArray(arrPM) THEN
		For intP = 0 To UBound(arrPM,2)
		 IF arrPM(1,intP)	= sadminId THEN
		 	blnMod = 1
		 	Exit For
		END IF
		Next
	END IF
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
    	parent.eappDetail.location.href = "confirmpayrequest.asp?iridx="+reportidx+"&ipridx="+payrequestidx+"&ias=<%=iauthstate%>&igbn=<%=igbn%>&iRM=<%=iRectMenu%>";
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
		<table width="100%" cellpadding="0" cellspacing="1" class="a" border="0">
 	<!----------------- 리스트 --------------------------->
		<form name="frmList" method="post" action="">
		<input type="hidden" name="iCP" value="<%=iCurrPage%>">
		<input type="hidden" name="iprs" value="<%=ipayrequeststate%>">
		<input type="hidden" name="ias" value="<%=iauthstate%>">
		<input type="hidden" name="iridx" value="<%=ireportidx%>">
		<input type="hidden" name="ipridx" value="<%=ipayrequestidx%>">
		<tr>
			<td>
				<!-- 검색 시작 -->
				<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
				<tr align="center" bgcolor="#FFFFFF" >
					<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
					<td align="left">
						* 요청서명 : <input type="text" name="reportname" id="reportname" value="<%= reportname %>" size="15" maxlength=150 class="text">
						<Br>
						* 요청금액 : <input type="text" name="reportprice" id="reportprice" value="<%= reportprice %>" size="8" maxlength=10 class="text">
						* 작성자 : <input type="text" name="username" id="username" value="<%= username %>" size="8" maxlength=10 class="text">
						<Br>
						* 부서 :
						<%= drawSelectBoxDepartment("department_id", department_id) %>
					</td>
					<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
						<input type="button" class="button_s" value="검색" onClick="frmsubmit('1');">
					</td>
				</tr>
				<tr align="center" bgcolor="#FFFFFF" >
					<td align="left">
						* 요청일 : <input type="text" id="termSdt" name="regdate" size="8" maxlength=10 value="<%= regdate %>" />
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
				</form>
				<!-- 검색 끝 -->
				<Br>
			</td>
	 	</tr>
		</form>
		<tr>
			<td height="25"><font color="#4E9FC6"><b>결제요청서> <%=fnGetMenu("FR",ipayrequeststate,iauthstate)%></b></font></td>
		</tr>
		<tr>
			<td>
				<table width="100%" align="left" cellpadding="5" cellspacing="1" class="a"   border="0" bgcolor="#cccccc">
				<tr bgcolor="<%= adminColor("tabletop") %>" align="center">
					<td rowspan="2">Idx</td>
					<td rowspan="2">결제요청서명</td>
					<td rowspan="2">결제요청금액</td>
					<td>결제요청일</td>
					<td>결제일</td>
					<td rowspan="2">작성자</td>
				</tr>
				<tr bgcolor="<%= adminColor("tabletop") %>" align="center">
					<td>결제예정일</td>
					<td>서류제출</td>
				</tr>
				<%IF isArray(arrList) THEN
					For intLoop = 0 To UBound(arrList,2)
				%>
				<tr id="t<%=arrList(0,intLoop)%>" bgcolor="#FFFFFF" align="center" onclick="jsConfirm(<%=arrList(8,intLoop)%>,<%=arrList(0,intLoop)%>);evalChangeColor('document.all.t<%=arrList(0,intLoop)%>','#CEF6EC','FFFFFF');" style="cursor:hand;">
						<td rowspan="2"><%=arrList(0,intLoop)%></td>
						<td rowspan="2">결제요청서(<%=arrList(14,intLoop)%>)<Br><font color="Gray"><%=arrList(12,intLoop)%></font></td>
						<td rowspan="2" nowrap><%IF arrList(2,intLoop)>0 THEN%><%=formatnumber(arrList(2,intLoop),0)%><%END IF%></td>
						<td nowrap><%=arrList(1,intLoop)%></td>
						<td nowrap><%IF arrList(4,intLoop) <> "" THEN%><%=formatdate(arrList(4,intLoop),"0000-00-00")%><%END IF%></td>
						<td nowrap rowspan="2"><%=arrList(11,intLoop)%></td>
				</tR>
				<tr id="t<%=arrList(0,intLoop)%>" bgcolor="#FFFFFF" align="center" onclick="jsConfirm(<%=arrList(8,intLoop)%>,<%=arrList(0,intLoop)%>);evalChangeColor('document.all.t<%=arrList(0,intLoop)%>','#CEF6EC','FFFFFF');" style="cursor:hand;">
					<td><%=arrList(3,intLoop)%></td>
					<td nowrap><%IF arrList(9,intLoop) THEN%>Y<%ELSE%>N<%END IF%><%IF blnMod = "1" THEN%>&nbsp;<a href="javascript:jsModTakeDoc('<%=arrList(0,intLoop)%>','<%=arrList(9,intLoop)%>');"><font color="blue">[수정]</font></a><%END IF%></td>
				</tr>
				<%
					Next
					ELSE
				%>
				<tr bgcolor="#FFFFFF">
					<td colspan="9" align="center">등록된 내역이 없습니다.</td>
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
					<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a"  >
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
	</td>
</tr>
</table>
</div>
<%IF ipayrequestidx >0 THEN%>
<script language="javascript">
	 //윈도우 로드시 해당 tr 색 변경
	 window.onload = jsOnSetColor;
   function jsOnSetColor(){
   	evalChangeColor('document.all.t<%=ipayrequestidx%>','#CEF6EC','#FFFFFF');
  }
</script>
<%END IF%>
<!-- 페이지 끝 -->
</body>
</html>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
'###############################################
' PageName : GncView.asp
' Discription : 모바일 GNB메뉴 관리
' History : 2018.01.11 원승현
'###############################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/appmanage/startupBannerCls.asp" -->
<!-- #include virtual="/partner/lib/adminHead.asp" -->
<!-- #include virtual="/lib/util/tenEncUtil.asp" -->
<%

	Dim vIdx, vMenuCode, vMenuName, vLinkURL, vStartDate, vEndDate, vRegDate, vLastUpDate, vAdminId
	Dim vLastAdminId, vOrderBy, vIsNew, vIsUsing, vAdminName, vLastAdminName
	Dim vGnbMenuName, strSql
	Dim vStartDateHour, vStartDateMinute, vStartDateSecond
	Dim vEndDateHour, vEndDateMinute, vEndDateSecond

	vIdx = requestCheckVar(request("idx"), 10)
	vMenuCode = requestCheckVar(request("MenuCode"), 20)
	vIsNew = False
	vIsUsing = False

	If vIdx <> "" Then
		'// idx값을 기준으로 GNB 메뉴값을 가져온다.
		strSql = " Select top 1 GM.idx, GM.MenuCode, GM.MenuName, GM.LinkURL, GM.StartDate, GM.EndDate, GM.RegDate "
		strSql = strSql & "	, GM.LastUpDate, GM.AdminId, GM.LastAdminId, GM.OrderBy, GM.IsNew, GM.IsUsing, "
		strSql = strSql & "		( "
		strSql = strSql & "			Select top 1 username From db_partner.[dbo].[tbl_user_tenbyten] Where userid = GM.AdminId "
		strSql = strSql & "		) as AdminName, "
		strSql = strSql & "		( "
		strSql = strSql & "			Select top 1 username From db_partner.[dbo].[tbl_user_tenbyten] Where userid = GM.LastAdminId "
		strSql = strSql & "		) as LastAdminName "
		strSql = strSql & " From db_sitemaster.[dbo].[tbl_GNBMenuManagement] GM "
		strSql = strSql & " Where idx = '"&vIdx&"' "
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		If Not(rsget.bof Or rsget.eof) Then
			'vCheck = True
			vIdx = rsget("idx")
			vMenuCode = rsget("MenuCode")
			vMenuName = rsget("MenuName")
			vLinkURL = rsget("LinkURL")
			vStartDate = rsget("StartDate")
			vEndDate = rsget("EndDate")
			vRegDate = rsget("RegDate")
			vLastUpDate = rsget("LastUpDate")
			vAdminId = rsget("AdminId")
			vLastAdminId = rsget("LastAdminId")
			vOrderBy = rsget("OrderBy")
			vIsNew = rsget("IsNew")
			vIsUsing = rsget("IsUsing")
			vAdminName = rsget("AdminName")
			vLastAdminName = rsget("LastAdminName")
		End If
		rsget.close
	End If

	Select Case Trim(vMenuCode)
		Case "SpecialA"
			vGnbMenuName = "GNBMenu1"
		Case "SpecialB"
			vGnbMenuName = "GNBMenu2"
		Case "SpecialC"
			vGnbMenuName = "GNBMenu3"
		Case "Class"
			vGnbMenuName = "클래스"
	End Select

	If Trim(vOrderBy)="" Then
		vOrderBy = 99
	End If

	If Trim(vStartDate) <> "" Then
		If Len(Hour(vStartDate)) < 2 Then
			vStartDateHour = "0"&Hour(vStartDate)
		Else
			vStartDateHour = Hour(vStartDate)
		End If
		If Len(Minute(vStartDate)) < 2 Then
			vStartDateMinute = "0"&Minute(vStartDate)
		Else
			vStartDateMinute = Minute(vStartDate)
		End If
		If Len(Second(vStartDate)) < 2 Then
			vStartDatesecond = "0"&Second(vStartDate)
		Else
			vStartDatesecond = Second(vStartDate)
		End If
	Else
		vStartDateHour = "00"
		vStartDateMinute = "00"
		vStartDatesecond = "00"
	End If

	If Trim(vEndDate) <> "" Then
		If Len(Hour(vEndDate)) < 2 Then
			vEndDateHour = "0"&Hour(vEndDate)
		Else
			vEndDateHour = Hour(vEndDate)
		End If
		If Len(Minute(vEndDate)) < 2 Then
			vEndDateMinute = "0"&Minute(vEndDate)
		Else
			vEndDateMinute = Minute(vEndDate)
		End If
		If Len(Second(vEndDate)) < 2 Then
			vEndDatesecond = "0"&Second(vEndDate)
		Else
			vEndDatesecond = Second(vEndDate)
		End If
	Else
		vEndDateHour = "23"
		vEndDateMinute = "59"
		vEndDatesecond = "59"
	End If


%>
<script type="text/javascript" src="/js/jquery.form.min.js"></script> 
<script type="text/javascript" src="/js/jsCal/js/jscal2.js"></script>
<script type="text/javascript" src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script type="text/javascript">

// 등록폼 확인 및 처리
function fnSubmit(frm) {
	if(frm.StartDate.value.length<10) {
		alert("시작일 입력해주세요.");
		frm.StartDate.focus();
		return false;
	}

	if(frm.EndDate.value<10) {
		alert("종료일 입력해주세요.");
		frm.EndDate.focus();
		return false;
	}

	if(!frm.MenuName.value) {
		alert("메뉴명을 입력해주세요.");
		frm.MenuName.focus();
		return false;
	}

	if(!frm.LinkURL.value.length) {
		alert("링크 URL을 입력해주세요.");
		frm.LinkURL.focus();
		return false;
	}

	if(confirm("입력하신 내용으로 등록하시겠습니까?")){
		frm.submit();
	}

}
</script>
</head>
<body>
<div class="popupWrap">
	<div class="popHead">
		<h1><img src="/images/partner/pop_admin_logo.gif" alt="10x10" /></h1>
		<p class="btnClose"><input type="image" src="/images/partner/pop_admin_btn_close.gif" alt="창닫기" onclick="window.close();" /></p>
	</div>
	<div class="popContent scrl" style="padding-top:20px;">
		<div class="contTit bgNone">
			<h2><%=vGnbMenuName%> 등록/수정</h2>
		</div>
		<div class="cont">
			<form name="frm1" action="doGNBReg.asp" method="post" style="margin:0px;">
			<input type="hidden" name="idx" value="<%=vidx%>">
			<input type="hidden" name="mode" value="<%=chkiif(vidx="" or isNull(vidx),"add","modi")%>">
			<input type="hidden" name="MenuCode" value="<%=vMenuCode%>">
				<table class="tbType1 writeTb" bgcolor="#FFFFFF">
					<tbody>
						<tr>
							<th width="12%">기간</th>
							<td height="30" style="padding-left:5px;">
								<input type="text" name="StartDate" value="<% If vStartDate <> "" Then response.write Left(vStartDate, 10) Else response.write "" %>" class="formTxt" id="termSdt" maxlength="10" style="width:100px" placeholder="시작일" readonly />
								<input type="text" name="StartDateSecond" class="formTxt" maxlength="12" style="width:100px" placeholder="시작일시분초" value="<%=vStartDateHour&":"&vStartDateMinute&":"&vStartDateSecond%>" />
								<input type="image" src="/images/admin_calendar.png" alt="달력으로 검색" id="ChkStart_trigger" onclick="return false;" />
								~
								<input type="text" name="EndDate" value="<% If vEndDate <> "" Then response.write Left(vEndDate, 10) Else response.write "" %>" class="formTxt" id="termEdt" maxlength="10" style="width:100px" placeholder="종료일" readonly />
								<input type="text" name="EndDateSecond" class="formTxt" maxlength="12" style="width:100px" placeholder="종료일일시분초" value="<%=vEndDateHour&":"&vEndDateMinute&":"&vEndDateSecond%>" />
								<input type="image" src="/images/admin_calendar.png" alt="달력으로 검색" id="ChkEnd_trigger" onclick="return false;" />
								<script type="text/javascript">
									var CAL_Start = new Calendar({
										inputField : "termSdt", trigger    : "ChkStart_trigger",
										onSelect: function() {
											var date = Calendar.intToDate(this.selection.get());
											CAL_End.args.min = date;
											CAL_End.redraw();
											this.hide();
										}, bottomBar: true, dateFormat: "%Y-%m-%d"
									});
									var CAL_End = new Calendar({
										inputField : "termEdt", trigger    : "ChkEnd_trigger",
										onSelect: function() {
											var date = Calendar.intToDate(this.selection.get());
											CAL_Start.args.max = date;
											CAL_Start.redraw();
											this.hide();
										}, bottomBar: true, dateFormat: "%Y-%m-%d"
									});
									var CAL_StartTxt = new Calendar({
										inputField : "termSdt", trigger    : "termSdt",
										onSelect: function() {
											var date = Calendar.intToDate(this.selection.get());
											CAL_End.args.min = date;
											CAL_End.redraw();
											this.hide();
										}, bottomBar: true, dateFormat: "%Y-%m-%d"
									});
									var CAL_EndTxt = new Calendar({
										inputField : "termEdt", trigger    : "termEdt",
										onSelect: function() {
											var date = Calendar.intToDate(this.selection.get());
											CAL_Start.args.max = date;
											CAL_Start.redraw();
											this.hide();
										}, bottomBar: true, dateFormat: "%Y-%m-%d"
									});
								</script>
							</td>
						</tr>
						<tr>
							<th>메뉴명</th>
							<td height="30" style="padding-left:5px;">
								<input type="text" name="MenuName" value="<%=vMenuName%>" class="formTxt" size="50" maxlength="5" />
							</td>
						</tr>
						<tr>
							<th>링크</th>
							<td height="30" style="padding-left:5px;">
								<p class="tMar05">주소 : <input type="text" name="LinkURL" value="<%=vLinkURL%>" class="formTxt" size="60" maxlength="180" /></p>
								<p class="tMar05">※ 아래 주소를 복사 후 이벤트코드만 입력해주시기 바랍니다.</p>
								<p class="tMar05">/gnbeventmain.asp?eventid=</p>
							</td>
						</tr>
						<tr>
							<th>정렬번호</th>
							<td height="30" style="padding-left:5px;">
								<p class="tMar05"><input type="text" name="OrderBy" value="<%=vOrderBy%>" class="formTxt" size="2" maxlength="10" /></p>
							</td>
						</tr>
						<tr>
							<th>New표시여부</th>
							<td height="30" style="padding-left:5px;">
								<label><input type="radio" name="IsNew" value="0" class="formCheck" <% If vIsNew = False Then %>checked<% End If %> /> 표시안함</label>
								<label><input type="radio" name="IsNew" value="1" class="formCheck" <% If vIsNew Then %>checked<% End If %> /> 표시함</label>
							</td>
						</tr>
						<tr>
							<th>사용여부</th>
							<td height="30" style="padding-left:5px;">
								<label><input type="radio" name="IsUsing" value="0" class="formCheck" <% If vIsUsing = False Then %>checked<% End If %> /> 사용안함</label>
								<label><input type="radio" name="IsUsing" value="1" class="formCheck" <% If vIsUsing Then %>checked<% End If %> /> 사용함</label>
							</td>
						</tr>
					</tboby>
				</table>

				<div class="tPad15 ct">
					<input type="button" value="취 소" onclick="if(confirm('작업을 취소하고 창을 닫겠습니까?')){self.close();}" class="btn3 btnDkGy" style="margin-right:30px;" />
					<input type="button" value="저 장" onclick="fnSubmit(this.form);" class="btn3 btnRd" />
				</div>
			</form>
		</div>
	</div>
</div>
</body>
</html>
<!-- #include virtual="/lib/db/dbclose.asp" -->
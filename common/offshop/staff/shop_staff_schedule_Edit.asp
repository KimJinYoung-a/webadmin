<%@ language=vbscript %>
<% option explicit %>
<%
	Response.AddHeader "Cache-Control","no-cache"
	Response.AddHeader "Expires","0"
	Response.AddHeader "Pragma","no-cache"
%>
<%
'####################################################
' Description :  오프라인 매장근무관리
' History : 2011.03.17 한용민 생성
'           2012.02.15 허진원- 미니달력 교체
'####################################################
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/md5.asp"-->
<!-- #include virtual="/common/incSessionAdminorShop.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/common/lib/commonbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/staff/staff_cls.asp"-->
<%
dim idx, sDt, sTm, eDt, eTm ,oAgitCal ,shopid ,empno
dim userid,username,posit_sn,part_sn,ChkStart,ChkEnd,etcComment
	idx = request("idx")
	shopid = request("shopid")

	if shopid = "" then
		resposne.write "<script>alert('매장이 지정되지 않았습니다'); self.close(); </script>"
		dbget.close()	:	response.end
	end if
	
'// 내용 접수
if idx<>"" then
	
	Set oAgitCal = new CAgitCalendar
		oAgitCal.frectidx = idx
		oAgitCal.read()
		
		if oAgitCal.ftotalcount >0 then
			userid			= oAgitCal.FOneItem.Fuserid
			username		= oAgitCal.FOneItem.Fusername
			posit_sn		= oAgitCal.FOneItem.Fposit_sn
			part_sn			= oAgitCal.FOneItem.Fpart_sn						
			ChkStart		= oAgitCal.FOneItem.FChkStart
			ChkEnd			= oAgitCal.FOneItem.FChkEnd			
			etcComment		= oAgitCal.FOneItem.FetcComment
			shopid			= oAgitCal.FOneItem.fshopid
			empno 			= oAgitCal.FOneItem.fempno
		
			sDt = left(ChkStart,10)
			eDt = left(ChkEnd,10)
			sTm = Num2Str(Hour(ChkStart),2,"0","R") & ":" & Num2Str(Minute(ChkStart),2,"0","R")& ":" & Num2Str(second(ChkStart),2,"0","R")
			eTm = Num2Str(Hour(ChkEnd),2,"0","R") & ":" & Num2Str(Minute(ChkEnd),2,"0","R")& ":" & Num2Str(second(ChkEnd),2,"0","R")
		end if
	Set oAgitCal = Nothing
end if

if sDt="" then sDt=date
if sTm="" then sTm="00:00:00"
if eDt="" then eDt=dateAdd("d",1,date)
if eTm="" then eTm="24:00:00"
	
part_sn = "18"	
%>

<script language="javascript1.2" type="text/javascript" src="/js/datetime.js"></script>
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />

<script language="javascript">

	// 등록 확인 및 처리
	function chk_form(form)	{
		if(form.shopid.value=="") {
			alert("매장을 선택해주세요");
			form.shopid.focus();
			return false;
		}

		if(form.chkCfm.value!="Y") {
			alert("검색 버튼을 눌러 직원을 검색해 주세요.");
			form.SearchText.focus();
			return false;
		}

		if(form.empno.value == "") {
			alert("사원번호를 입력 하세요");
			form.empno.focus();
			return false;
		}

		if(form.posit_sn.value == "") {
			alert("직급을 선택해주세요.");
			form.posit_sn.focus();
			return false;
		}

		if(form.username.value == "") {
			alert("이름을 입력해주세요.");
			form.username.focus();
			return false;
		}

		if(form.part_sn.value == "") {
			alert("소속부서를 선택해주세요.");
			form.part_sn.focus();
			return false;
		}

		//if(getDayInterval(toDate(form.ChkStart.value), toDate('<%=date%>'))>0) {
		//	alert("지나간 날짜는 등록하실 수 없습니다. 날짜를 확인해주세요.");
		//	return false;
		//}

		if(getDayInterval(toDate(form.ChkStart.value), toDate(form.ChkEnd.value))<0) {
			alert("기간이 잘못되어 있습니다. 날짜를 확인해주세요.");
			return false;
		}

		if(confirm(form.ChkStart.value +"~"+form.ChkEnd.value +"기간에("+form.uTerm.value+"일) " + form.username.value + "님을 등록하시겠습니까?"))	{
		return true;
		}
		return false;
	}

	//이용 기간 확인 및 박수 자동입력
	function chkTerm() {
		var frm = document.frm;
		
		var startday = frm.ChkStart;
		var endday = frm.ChkEnd;
	
		var startdate = toDate(startday.value);
		var enddate = toDate(endday.value);
	
		if ((startday.value == "") || (endday.value == "")) {
			alert("기간을 입력해주십시요.");
			return;
		}
	
		if (getDayInterval(startdate, enddate) < 0) {
			//alert("잘못된 기간입니다.");
			//return;
		}
	
		frm.uTerm.value = getDayInterval(startdate, enddate)+1;
	}

	//직원 아이디 검사 및 관련내용 자동입력
	function chkTenMember() {
		var SearchType;
		var SearchText;
		var shopid;
		
		if(frm.SearchType.value == '') {
			alert("검색하실 조건을 선택하세요");
			frm.SearchType.focus();
			return;
		}

		if(frm.SearchText.value == '') {
			alert("검색하실 값을 입력 하세요");
			frm.SearchText.focus();
			return;
		}		

		if(frm.shopid.value == '') {
			alert("선택된 매장이 없습니다");			
			return;
		}	
		
		SearchType = frm.SearchType.value;
		SearchText = frm.SearchText.value;
		shopid = frm.shopid.value;
		document.getElementById("ifmProc").src="/common/offshop/staff/actionTenUser.asp?SearchType="+SearchType+"&SearchText="+SearchText+"&shopid="+shopid;
	}

	//삭체처리
	function delBook() {
		if(confirm("본 예약내역을 삭제하시겠습니까?"))	{
			frm.mode.value = "del";
			frm.submit();
		}
	}

</script>

<table width="100%" border="0" cellpadding="5" cellspacing="0" class="a">
<form name="frm" method="POST" action="/common/offshop/staff/shop_staff_schedule_Process.asp" onsubmit="return chk_form(this)">
<input type="hidden" name="mode" value="<%=chkIIF(idx="","add","modi")%>">
<input type="hidden" name="idx" value="<%=idx%>">
<tr>
	<td>
		<table width="100%" border="0" cellpadding="5" cellspacing="1" class="a" bgcolor="#909090">
		<tr bgcolor="#FFFFFF">
			<td width="120" bgcolor="<%=adminColor("sky")%>" align="center"><b>매장</b></td>
			<td>				
				<%= shopid %><input type="hidden" name="shopid" value="<%=shopid%>">
			</td>
		</tr>
		<tr bgcolor="#FFFFFF">
			<td width="120" bgcolor="<%=adminColor("sky")%>" align="center"><b>사원검색</b></td>
			<td>
				<select name="SearchType">
					<option value="2">이름</option>
					<option value="1">아이디</option>					
					<option value="3">사번</option>
				</select>				
				<input type="text" name="SearchText" size="20" class="text">
				<input type="button" value="검색" class="button_s" style="width:55px;text-align:center;" onclick="chkTenMember()">				
				<input type="hidden" name="chkCfm" value="<%=chkIIF(idx="","N","Y")%>">
			</td>
		</tr>
		<tr bgcolor="#FFFFFF">
			<td width="120" bgcolor="<%=adminColor("sky")%>" align="center"><b>사원번호</b></td>
			<td>				
				<input type="text" name="empno" size="16" class="text" value="<%=empno%>">
			</td>
		</tr>
		<tr bgcolor="#FFFFFF">
			<td width="120" bgcolor="<%=adminColor("sky")%>" align="center"><b>아이디</b></td>
			<td>				
				<input type="text" name="userid" size="16" class="text" value="<%=userid%>">
			</td>
		</tr>		
		<tr bgcolor="#FFFFFF">
			<td width="120" bgcolor="<%=adminColor("sky")%>" align="center"><b>직급/이름</b></td>
			<td>
				<%=printPositOption("posit_sn", posit_sn)%>
				<input type="text" name="username" size="16" class="text" value="<%=username%>">
			</td>
		</tr>
		<tr bgcolor="#FFFFFF">
			<td width="120" bgcolor="<%=adminColor("sky")%>" align="center"><b>소속부서</b></td>
			<td><%=printPartOption("part_sn", part_sn)%></td>
		</tr>
		<tr bgcolor="#FFFFFF">
			<td width="120" bgcolor="<%=adminColor("sky")%>" align="center"><b>기간</b></td>
			<td style="line-height:18px;">
				<input id="ChkStart" name="ChkStart" value="<%=sDt%>" class="text" size="10" maxlength="10" /><img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="ChkStart_trigger" border="0" style="cursor:pointer" align="absmiddle" />
		    	<input type="text" name="ChkSTime" size="8" maxlength="8" class="text" value="<%=sTm%>">
		    	~
				<input id="ChkEnd" name="ChkEnd" value="<%=eDt%>" class="text" size="10" maxlength="10" /><img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="ChkEnd_trigger" border="0" style="cursor:pointer" align="absmiddle" />
		    	<input type="text" name="ChkETime" size="8" maxlength="8" class="text" value="<%=eTm%>">
		    	<font color=gray>(<input type="text" name="uTerm" readonly class="text" value="<%=DateDiff("d",sDt,eDt)+1%>" style="text-align:right; width:20px; border:0px; color:gray;">일)</font>
				<script language="javascript">
					var CAL_Start = new Calendar({
						inputField : "ChkStart", trigger    : "ChkStart_trigger",
						onSelect: function() {
							var date = Calendar.intToDate(this.selection.get());
							CAL_End.args.min = date;
							CAL_End.redraw();
							this.hide();
						}, bottomBar: true, dateFormat: "%Y-%m-%d"
					});
					var CAL_End = new Calendar({
						inputField : "ChkEnd", trigger    : "ChkEnd_trigger",
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
		<tr bgcolor="#FFFFFF">
			<td width="120" bgcolor="<%=adminColor("sky")%>" align="center"><b>비고</b></td>
			<td><textarea name="etcComment" class="textarea" style="width:100%; height:50px;"><%=etcComment%></textarea></td>
		</tr>
		<tr bgcolor="#FFFFFF">
			<td colspan="2" align="center">
				<input type="submit" value="등 록" class="button" style="width:60px;text-align:center;">
				<% if idx<>"" then %>
					<input type="button" value="삭 제" class="button" style="width:60px;text-align:center;" onclick="delBook()">
				<% end if %>
			</td>
		</tr>
		</table>
	</td>
</tr>
</form>
</table>

<iframe id="ifmProc" src="" width=0 height=0 frameborder="0"></iframe>
<!-- #include virtual="/common/lib/commonbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
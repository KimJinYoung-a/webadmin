<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : cs센터 고객조회
' History : 2009.04.17 이상구 생성
'           2023.10.30 한용민 수정(휴면계정정보표기. 휴면계정->일반계정 전환 로직 생성)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/checkAllowIPWithLog.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/member/customercls.asp"-->
<!-- #include virtual="/lib/classes/member/offlinecustomercls.asp"-->
<!-- #include virtual="/lib/classes/mileage/sp_pointcls.asp" -->
<!-- #include virtual="/lib/classes/mileage/sp_mileage_logcls.asp" -->
<!-- #include virtual="/lib/classes/cscenter/cs_couponcls.asp" -->
<!-- #include virtual="/lib/classes/cscenter/sp_itemcouponcls.asp" -->
<!-- #include virtual="/lib/classes/event/eventPrizeCls.asp"-->
<!-- #include virtual="/admin/eventmanage/common/event_function.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<%
dim userid, userseq, i, buf, haveofflineaccount, haveonlineaccount, maxModifyDate, OUserInfo, OOfflineUserInfo, myMileage
dim oitemcoupon, ocscoupon, myOffMileage, oExpireMile, clsEPrize, arrList, iDelCnt, total_event_count, total_before_verify_count
dim issameusercell, issameusermail, sqlStr, issameuserphone, snsGubunList, snsGubun
	userid = requestCheckvar(request("userid"),32)
	userseq = requestCheckvar(request("userseq"),32)

if ((userid = "") and (userseq = "")) then
    'response.write "<script>alert('잘못된 접속입니다.'); history.back();</script>"
    'dbget.close()	:	response.End
end if

if (userid <> "") then
	haveonlineaccount = "Y"

	set OUserInfo = new CUserInfo
		OUserInfo.FRectUserID = userid
		OUserInfo.GetUserInfo

	if OUserInfo.FTotalCount<1 then
		response.write "<script type='text/javascript'>"
		response.write "	alert('회원정보가 존재하지 않습니다.');"
		response.write "	self.close();"
		response.write "</script>"
		response.write "회원정보가 존재하지 않습니다."
		dbget.close() : response.end
	end if

	maxModifyDate = OUserInfo.GetUserMaxModifyDate()

	if OUserInfo.Fitemlist(0).Fuserdiv = "05" then
		snsGubunList = OUserInfo.GetSNSUserJoinPathList
		if isArray(snsGubunList) then
			for i=0 to UBound(snsGubunList,2)
				snsGubun = snsGubun & chkIIF(snsGubun<>""," / ","") & GetSNSJoinTypeName(snsGubunList(0,i))
			Next
		end if
	end if	

	set OOfflineUserInfo = new COfflineUserInfo
		OOfflineUserInfo.FRectUserID = userid
		OOfflineUserInfo.GetUserInfo

	if (OOfflineUserInfo.FTotalCount > 0) then

		haveofflineaccount = "Y"
		userseq = OOfflineUserInfo.Fitemlist(0).FUserSeq

	else
		haveofflineaccount = "N"

		'redim OOfflineUserInfo.FItemList(1)
		'set OOfflineUserInfo.FItemList(i) = new COfflineUserInfoItem
	end if
else
	haveofflineaccount = "Y"

	set OOfflineUserInfo = new COfflineUserInfo
		OOfflineUserInfo.FRectUserSeq = Cint(userseq)
		OOfflineUserInfo.GetUserInfo

	if ((OOfflineUserInfo.FTotalCount > 0) and (OOfflineUserInfo.Fitemlist(0).FUserID <> "")) then

		haveonlineaccount = "Y"
		userid = OOfflineUserInfo.Fitemlist(0).FUserID

		OUserInfo.FRectUserID = userid
		OUserInfo.GetUserInfo

	else
		haveonlineaccount = "N"

		'redim preserve OUserInfo.FItemList(1)
		'set OUserInfo.FItemList(i) = new CUserInfoItem
	end if
end if

if (haveonlineaccount = "Y") then
	set myMileage = new TenPoint
		myMileage.FRectUserID = userid

		if (userid <> "") then
			myMileage.getTotalMileage
		end if

	set myOffMileage = new TenPoint
		myOffMileage.FGubun = "my10x10"
		myOffMileage.FRectUserID = userid

	if (userid <> "") then
	    myOffMileage.getOffShopMileagePop
	end if

	''만료예정  마일리지
	set oExpireMile = new CMileageLog
		oExpireMile.FRectUserid = userid
		oExpireMile.FRectExpireDate = Left(CStr(now()),4) + "-12-31"

	if (userid<>"") then
	    oExpireMile.getNextExpireMileageSum
	end if
end if

'상품쿠폰
set oitemcoupon = new CUserItemCoupon
	oitemcoupon.FRectUserID = userid
	oitemcoupon.FRectAvailableYN = "Y"
	oitemcoupon.FRectDeleteYN = "Y"
	oitemcoupon.FPageSize = 200
	oitemcoupon.FCurrPage = 1
	oitemcoupon.GetCouponList

'보너스쿠폰
set ocscoupon = New CCSCenterCoupon
	ocscoupon.FRectExcludeUnavailable = "Y"
	ocscoupon.FRectExcludeDelete = "Y"
	ocscoupon.FRectUserID = userid
	ocscoupon.GetCSCenterCouponList

'당첨
set clsEPrize = new CEventPrize
	clsEPrize.FSUserid = userid
	clsEPrize.FPSize = 20
	clsEPrize.FCPage = 1
	arrList = clsEPrize.fnGetPrizeList

total_event_count = clsEPrize.FTotCnt

clsEPrize.FEPStatus = "0"
arrList = clsEPrize.fnGetPrizeList
total_before_verify_count = clsEPrize.FTotCnt

if ((haveonlineaccount = "Y") and (haveofflineaccount = "Y")) then

	sqlStr = "insert into db_cs.dbo.tbl_cs_usersearch_Log(customeruserid, offcustomerseq, adminuserid, searchip)"
	sqlStr = sqlStr + " values('" & userid & "', " & userseq & ", '" & session("ssBctId") & "', '" & Request.ServerVariables("REMOTE_ADDR") & "') "

elseif (haveonlineaccount = "Y") then
	sqlStr = "insert into db_cs.dbo.tbl_cs_usersearch_Log(customeruserid, adminuserid, searchip)"
	sqlStr = sqlStr + " values('" & userid & "', '" & session("ssBctId") & "', '" & Request.ServerVariables("REMOTE_ADDR") & "') "
else
	sqlStr = "insert into db_cs.dbo.tbl_cs_usersearch_Log(offcustomerseq, adminuserid, searchip)"
	sqlStr = sqlStr + " values(" & userseq & ", '" & session("ssBctId") & "', '" & Request.ServerVariables("REMOTE_ADDR") & "') "
end if
rsget.CursorLocation = adUseClient
rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly

issameusercell = "N"
issameuserphone = "N"
issameusermail = "N"
if (haveonlineaccount = "Y") and (haveofflineaccount = "Y") then

	if (OUserInfo.FItemList(0).Fusercell = OOfflineUserInfo.FItemList(0).Fusercell) then
		issameusercell = "Y"
	end if
	if (OUserInfo.FItemList(0).Fuserphone = OOfflineUserInfo.FItemList(0).Fuserphone) then
		issameuserphone = "Y"
	end if
	if (OUserInfo.FItemList(0).FUsermail = OOfflineUserInfo.FItemList(0).FUsermail) then
		issameusermail = "Y"
	end if

end if

%>
<script type="text/javascript">

function popYearExpireMileList(yyyymmdd,userid){
    var popwin = window.open('/cscenter/mileage/popAdminExpireMileSummary.asp?yyyymmdd=' + yyyymmdd + '&userid=' + userid,'popAdminExpireMileSummary','width=660,height=500,scrollbars=yes,resizable=yes');
    popwin.focus();
}

function popMileList(userid){
    var popwin = window.open('/cscenter/mileage/cs_mileage.asp?menupos=964&userid=' + userid,'popMileList','width=1000,height=800,scrollbars=yes,resizable=yes');
    popwin.focus();
}

function popCouponList(userid){
    var popwin = window.open('/cscenter/coupon/cs_coupon.asp?menupos=965&userid=' + userid,'popCouponList','width=1000,height=800,scrollbars=yes,resizable=yes');
    popwin.focus();
}

function popEventList(userid){
    var popwin = window.open('/admin/eventmanage/event/eventprize_list.asp?menupos=1056&searchUserid=' + userid,'popEventList','width=1000,height=800,scrollbars=yes,resizable=yes');
    popwin.focus();
}

// 휴대폰번호 삭제
function DelOnUserCellPhone(frm) {

	<% if issameusercell = "Y" then %>
		alert("온/오프라인 의 고객정보에 동일한 핸드폰 번호가 있습니다.\n\n두 핸드폰 번호를 모두 삭제합니다.(CS메모에 내역저장)");
	<% end if %>

	if (confirm('핸드폰 번호를 삭제합니다.(CS메모에 내역저장)\n\n진행하시겠습니까?') == true) {
		frm.mode.value = "delonusercell";
		frm.submit();
	}
}

// 일반전화번호 삭제
function DelOnUserPhone(frm) {

	<% if issameuserphone = "Y" then %>
		alert("온/오프라인 의 고객정보에 동일한 전화번호가 있습니다.\n\n두 전화번호를 모두 삭제합니다.(CS메모에 내역저장)");
	<% end if %>

	if (confirm('전화번호를 삭제합니다.(CS메모에 내역저장)\n\n진행하시겠습니까?') == true) {
		frm.mode.value = "delonuserphone";
		frm.submit();
	}
}

function DelOnUserMail(frm) {

	<% if issameusermail = "Y" then %>
		alert("온/오프라인 의 고객정보에 동일한 이메일주소가 있습니다.\n\n두 이메일주소를 모두 삭제합니다.(CS메모에 내역저장)");
	<% end if %>

	if (confirm('이메일주소를 삭제합니다.(CS메모에 내역저장)\n\n진행하시겠습니까?') == true) {
		frm.mode.value = "delonusermail";
		frm.submit();
	}
}

function ResetUserPass(frm) {
	if (confirm("\n\n주의!!!!\n\n임시 비밀번호를 생성합니다.\n\n임시비밀번호는 자동으로 발송되지 않으며 CS메모에만 기록됩니다.\n(별도 고객안내 필요)\n\n진행하시겠습니까?") == true) {
		frm.mode.value = "resetUserPass";
		frm.submit();
	}
}

function DelOffUserCellPhone(frm) {

	<% if issameusercell = "Y" then %>
		alert("온/오프라인 의 고객정보에 동일한 핸드폰 번호가 있습니다.\n\n두 핸드폰 번호를 모두 삭제합니다.(CS메모에 내역저장)");
	<% end if %>

	if (confirm('핸드폰 번호를 삭제합니다.\n\n진행하시겠습니까?') == true) {
		frm.mode.value = "deloffusercell";
		frm.submit();
	}
}

function SetUserDivTo01(frm) {
	if (confirm("일반회원으로 전환합니다.\n\n진행하시겠습니까?") == true) {
		frm.mode.value = "setuserdivto01";
		frm.submit();
	}
}

</script>

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
		고객정보 조회내역은 별도로 기록됩니다.
	</td>
	<td align="right">		
	</td>
</tr>
</table>
<!-- 액션 끝 -->

<br>

<form name="frm" method="post" action="/cscenter/member/domodifyuserinfo.asp" onsubmit="return false;" style="margin:0px;">
<input type="hidden" name="mode" value="modifyuserinfo">
<input type="hidden" name="userid" value="<%= userid %>">
<input type="hidden" name="userseq" value="<%= userseq %>">
<input type="hidden" name="haveonlineaccount" value="<%= haveonlineaccount %>">
<input type="hidden" name="haveofflineaccount" value="<%= haveofflineaccount %>">
<input type="hidden" name="issameusercell" value="<%= issameusercell %>">
<input type="hidden" name="issameuserphone" value="<%= issameuserphone %>">
<input type="hidden" name="issameusermail" value="<%= issameusermail %>">
<table width="100%" border="0" align="center" class="a" cellpadding="0" cellspacing="1" bgcolor="#BABABA">
<tr align="left">
	<td colspan=6 bgcolor="#FFFFFF">기본정보 [최종수정일 : <%= maxModifyDate %>]</td>
</tr>

<% if (haveonlineaccount = "Y") then %>
	<tr align="left">
		<td height="30" width="15%" bgcolor="#DDDDFF">아이디 :</td>
		<td bgcolor="#FFFFFF" colspan="3" width="35%" >
			<%= userid %>
			<% if Not(OUserInfo.Fitemlist(0).Fuserdiv = "05") then %>
			&nbsp; <input type="button" class="button" value="임시비밀번호 생성" onClick="ResetUserPass(frm)">
			<% else %>
			<br /><span style="color:#A55;font-size:9pt;">(회원전환 후 비밀번호를 생성할 수 있습니다.)</span>
			<% end if %>
		</td>
		<td height="30" width="15%" bgcolor="#DDDDFF">고객명 :</td>
		<td bgcolor="#FFFFFF" colspan="3" width="35%" >
			<%= OUserInfo.Fitemlist(0).FUserName %>
		</td>
	</tr>
	<tr align="left">
		<td height="30" width="15%" bgcolor="#DDDDFF">회원가입방식</td>
		<td bgcolor="#FFFFFF" colspan="3">
			<%
			if (OUserInfo.Fitemlist(0).Fuserdiv = "01") then
				response.write "일반회원"
			elseif (OUserInfo.Fitemlist(0).Fuserdiv = "05") then
				response.write "SNS가입회원 (" & snsGubun & ")&nbsp; <input type='button' class='button' value='일반회원전환' onclick='SetUserDivTo01(frm)'>"
			elseif (OUserInfo.Fitemlist(0).Fuserdiv = "96") then
				response.write "차단 기타 회원 (정지회원)"
			end if
			%>
		</td>
		<td height="30" width="15%" bgcolor="#DDDDFF">생일 :</td>
		<td bgcolor="#FFFFFF" colspan="3"><%= OUserInfo.Fitemlist(0).Fbirthday %> [<% if (OUserInfo.Fitemlist(0).Fissolar = "Y") then %>양력<% else %>음력<% end if %>]</td>
	</tr>
<% else %>
	<tr align="left">
		<td height="30" width="15%" bgcolor="#DDDDFF">아이디 :</td>
		<td bgcolor="#FFFFFF" colspan="3" width="35%" >
			<%= userid %>
			&nbsp;
			<input type="button" class="button" value="비밀번호 초기화" onClick="ResetUserPass(frm)">
		</td>
		<td height="30" width="15%" bgcolor="#DDDDFF">고객명 :</td>
		<td bgcolor="#FFFFFF" colspan="3" width="35%" >
			<%= OOfflineUserInfo.Fitemlist(0).FUserName %>
		</td>
	</tr>
	<tr align="left">
		<td height="30" width="15%" bgcolor="#DDDDFF">주민번호 :</td>
		<td bgcolor="#FFFFFF" colspan="3">
			<%
			if (Len(OOfflineUserInfo.FItemList(i).FJuminNo) > 6) then
				response.write Left(OOfflineUserInfo.FItemList(i).FJuminNo, (Len(OOfflineUserInfo.FItemList(i).FJuminNo) - 6)) & "******"
			else
				response.write OOfflineUserInfo.FItemList(i).FJuminNo
			end if
			%>
		</td>
		<td height="30" width="15%" bgcolor="#DDDDFF">생일 :</td>
		<td bgcolor="#FFFFFF" colspan="3"></td>
	</tr>
<% end if %>
</table>

<% if (haveonlineaccount = "Y") then %>
<br>
<table width="100%" border="0" align="center" class="a" cellpadding="0" cellspacing="1" bgcolor="#BABABA">
<tr align="left">
	<td colspan=7 bgcolor="#FFFFFF">연락처 - 온라인</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">전화번호 :</td>
	<td bgcolor="#FFFFFF" colspan="3" width="35%">
		<%= OUserInfo.FItemList(0).Fuserphone %>
		&nbsp;
		<% if (OUserInfo.FItemList(0).Fuserphone <> "") and (Not IsNull(OUserInfo.FItemList(0).Fuserphone)) then %>
			<input type="button" class="button" value=" 번호삭제 " onClick="DelOnUserPhone(frm)">
		<% end if %>
	</td>
	<td height="30" width="15%" bgcolor="#DDDDFF">핸드폰번호 :</td>
	<td bgcolor="#FFFFFF" colspan="3" width="35%">
		<%= OUserInfo.FItemList(0).Fusercell %>
		&nbsp;
		<% if (OUserInfo.FItemList(0).Fusercell <> "") and (Not IsNull(OUserInfo.FItemList(0).Fusercell)) then %>
			<input type="button" class="button" value=" 번호삭제 " onClick="DelOnUserCellPhone(frm)">
		<% end if %>
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">주소 :</td>
	<td bgcolor="#FFFFFF" colspan="3">
		[<%= OUserInfo.FItemList(0).Fzipcode %>] <%= OUserInfo.FItemList(0).Faddress1 %> <%= OUserInfo.FItemList(0).Faddress2 %>
	</td>
	<td height="30" width="15%" bgcolor="#DDDDFF">이메일 :</td>
	<td bgcolor="#FFFFFF" colspan="3">
		<%= OUserInfo.Fitemlist(0).FUsermail %>
		&nbsp;
		<% if (OUserInfo.FItemList(0).FUsermail <> "") and (Not IsNull(OUserInfo.FItemList(0).FUsermail)) then %>
		<input type="button" class="button" value=" 이메일삭제 " onClick="DelOnUserMail(frm)">
		<% end if %>
	</td>
</tr>
</table>
<% end if %>

<% if (haveofflineaccount = "Y") then %>
<br>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<tr align="left">
	<td colspan=8 bgcolor="#FFFFFF">연락처 - 오프라인</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">전화번호 :</td>
	<td bgcolor="#FFFFFF" colspan="3" width="35%"><%= OOfflineUserInfo.FItemList(0).Fuserphone %></td>
	<td height="30" width="15%" bgcolor="#DDDDFF">핸드폰번호 :</td>
	<td bgcolor="#FFFFFF" colspan="3" width="35%">
		<%= OOfflineUserInfo.FItemList(0).Fusercell %>
		&nbsp;
		<% if (OOfflineUserInfo.FItemList(0).Fusercell <> "") and (Not IsNull(OOfflineUserInfo.FItemList(0).Fusercell)) then %>
		<input type="button" class="button" value=" 번호삭제 " onClick="DelOffUserCellPhone(frm)">
		<% end if %>
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">주소 :</td>
	<td bgcolor="#FFFFFF" colspan="3">
		[<%= OOfflineUserInfo.FItemList(0).Fzipcode %>] <%= OOfflineUserInfo.FItemList(0).Faddress1 %> <%= OOfflineUserInfo.FItemList(0).Faddress2 %>
	</td>
	<td height="30" width="15%" bgcolor="#DDDDFF">이메일 :</td>
	<td bgcolor="#FFFFFF" colspan="3">
		<%= OOfflineUserInfo.Fitemlist(0).FUsermail %>
	</td>
</tr>
</table>
<% end if %>

<% if (haveonlineaccount = "Y") then %>
<br>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr align="left">
	<td colspan=6 bgcolor="#FFFFFF">
		마일리지 >>>>>>> <a href="javascript:popMileList('<%= userid %>')">상세내역보기</a>
	</td>
</tr>
<tr align="center" bgcolor="#DDDDFF">
	<td height=25>구분</td>
	<td>현재마일리지</td>
	<td>보너스 마일리지</td>
	<td>구매 마일리지</td>
	<td>사용한 마일리지</td>
	<td>소멸된 마일리지</td>
</tr>

	<% if (userid <> "") then %>
		<tr align="center" bgcolor="#FFFFFF">
			<td height=25>온라인</td>
			<td><strong><%=FormatNumber(myMileage.FTotalMileage,0) %></strong></td>
			<td><%=FormatNumber(myMileage.FBonusMileage,0) %></td>
			<td><%=FormatNumber(myMileage.FTotJumunmileage + myMileage.FAcademymileage,0) %></td>
			<td><%=FormatNumber(myMileage.FSpendMileage*-1,0) %></font></td>
			<td><%=FormatNumber(myMileage.FrealExpiredMileage*-1,0) %></font></td>
		</tr>
		<tr align="center" bgcolor="#FFFFFF">
			<td height=25>오프라인</td>
			<td><strong><%=FormatNumber(myOffMileage.FOffShopMileage,0) %></strong></td>
			<td colspan=4></td>
		</tr>
		<tr align="center" bgcolor="#FFFFFF">
			<td height=25>소멸 대상 마일리지</td>
			<td><a href="javascript:popYearExpireMileList('<%= oExpireMile.FRectExpireDate %>','<%= userid %>');"><%= FormatNumber(oExpireMile.FOneItem.getMayExpireTotal,0) %></a></td>
			<td colspan=4 align=left> &nbsp;&nbsp;<a href="javascript:popYearExpireMileList('<%= oExpireMile.FRectExpireDate %>','<%= userid %>');">* 소멸일자 : <%= Left(CStr(now()),4) + "-12-31" %></a></td>
		</tr>

			<% if (myMileage.FOldJumunmileage>0) then %>
		<tr align="center" bgcolor="#FFFFFF">
			<td height=25>6개월이전 적립합계</td>
			<td><%= FormatNumber(myMileage.FOldJumunmileage,0) %></td>
			<td colspan=4 align=left></td>
		</tr>
			<% else %>
		<tr align="center" bgcolor="#FFFFFF">
			<td height=25>6개월이전 적립합계</td>
			<td>없음</td>
			<td colspan=4 align=left></td>
		</tr>
			<% end if %>
			<% if (myMileage.FAcademyMileage>0) then %>
		<tr align="center" bgcolor="#FFFFFF">
			<td height=25>아카데미 주문적립</td>
			<td><%= FormatNumber(myMileage.FAcademyMileage,0) %></td>
			<td colspan=4></td>
		</tr>
			<% else %>
		<tr align="center" bgcolor="#FFFFFF">
			<td height=25>아카데미 주문적립</td>
			<td>없음</td>
			<td colspan=4></td>
		</tr>
			<% end if %>
	<% else %>
		<tr align="center" bgcolor="#FFFFFF">
			<td>온라인</td>
			<td>-</td>
			<td>-</td>
			<td>-</td>
			<td>-</td>
			<td>-</td>
		</tr>
	<% end if %>
</table>
<% end if %>

<% if (haveonlineaccount = "Y") then %>
<br>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<tr align="left">
	<td colspan=8 bgcolor="#FFFFFF">
		쿠폰 >>>>>>> <a href="javascript:popCouponList('<%= userid %>')">상세내역보기</a>
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">사용가능한 상품쿠폰 :</td>
	<td bgcolor="#FFFFFF" colspan="3" width="35%"><%= oitemcoupon.FTotalCount %></td>
	<td height="30" width="15%" bgcolor="#DDDDFF">사용가능한 보너스쿠폰 :</td>
	<td bgcolor="#FFFFFF" colspan="3" width="35%"><%= ocscoupon.FResultCount %></td>
</tr>
</table>
<% end if %>

<% if (haveonlineaccount = "Y") then %>
<br>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<tr align="left">
	<td colspan=8 bgcolor="#FFFFFF">
		당첨 >>>>>>> <a href="javascript:popEventList('<%= userid %>')">상세내역보기</a>
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">전체 당첨건수 :</td>
	<td bgcolor="#FFFFFF" colspan="3" width="35%"><%= total_event_count %></td>
	<td height="30" width="15%" bgcolor="#DDDDFF">당첨후 미확인 건수 :</td>
	<td bgcolor="#FFFFFF" colspan="3" width="35%"><%= total_before_verify_count %></td>
</tr>
</table>
<% end if %>

<br>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<tr align="left">
	<td colspan=8 bgcolor="#FFFFFF">
		이메일 수신여부
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">텐바이텐 :</td>
	<td bgcolor="#FFFFFF" colspan="3" width="35%">
		<% if (haveonlineaccount = "Y") then %>
			<table class="a" border="0" cellspacing="0" cellpadding="0">
			<tr>
				<td style="padding-bottom:2px;"><input type="radio" name="mail10x10" value="Y" <% if (OUserInfo.Fitemlist(0).Fmail10x10 = "Y") then %>checked<% end if %>></td>
				<td style="padding-left:2px;">예</td>
				<td style="padding:0 0 2px 15px;"><input type="radio"  name="mail10x10" value="N"  <% if (OUserInfo.Fitemlist(0).Fmail10x10 = "N") then %>checked<% end if %>></td>
				<td style="padding-left:2px;">아니오</td>
			</tr>
			</table>
		<% else %>
			계정없음
		<% end if %>
	</td>
	<td height="30" width="15%" bgcolor="#DDDDFF">핑거스 :</td>
	<td bgcolor="#FFFFFF" colspan="3" width="35%">
		<% if (haveonlineaccount = "Y") then %>
			<table class="a" border="0" cellspacing="0" cellpadding="0">
			<tr>
				<td style="padding-bottom:2px;"><input type="radio" name="mailfinger" value="Y" <% if (OUserInfo.Fitemlist(0).Fmailfinger = "Y") then %>checked<% end if %>></td>
				<td style="padding-left:2px;">예</td>
				<td style="padding:0 0 2px 15px;"><input type="radio"  name="mailfinger" value="N"  <% if (OUserInfo.Fitemlist(0).Fmailfinger = "N") then %>checked<% end if %>></td>
				<td style="padding-left:2px;">아니오</td>
			</tr>
			</table>
		<% else %>
			계정없음
		<% end if %>
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">포인트 :</td>
	<td bgcolor="#FFFFFF" colspan="3">
		<% if (haveofflineaccount = "Y") then %>
			<table class="a" border="0" cellspacing="0" cellpadding="0">
			<tr>
				<td style="padding-bottom:2px;"><input type="radio" name="offlinemail" value="Y" <% if (OOfflineUserInfo.Fitemlist(0).Fmail = "Y") then %>checked<% end if %>></td>
				<td style="padding-left:2px;">예</td>
				<td style="padding:0 0 2px 15px;"><input type="radio"  name="offlinemail" value="N"  <% if (OOfflineUserInfo.Fitemlist(0).Fmail = "N") then %>checked<% end if %>></td>
				<td style="padding-left:2px;">아니오</td>
			</tr>
			</table>
		<% else %>
			계정없음
		<%end if %>
	</td>
	<td height="30" width="15%" bgcolor="#DDDDFF"></td>
	<td bgcolor="#FFFFFF" colspan="3"></td>
</tr>
</table>
<br>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<tr align="left">
	<td colspan=8 bgcolor="#FFFFFF">
		SMS 수신여부
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">텐바이텐 :</td>
	<td bgcolor="#FFFFFF" colspan="3" width="35%">
		<% if (haveonlineaccount = "Y") then %>
			<table class="a" border="0" cellspacing="0" cellpadding="0">
			<tr>
			<td style="padding-bottom:2px;"><input type="radio" name="sms10x10" value="Y" <% if (OUserInfo.Fitemlist(0).Fsms10x10 = "Y") then %>checked<% end if %>></td>
			<td style="padding-left:2px;">예</td>
			<td style="padding:0 0 2px 15px;"><input type="radio"  name="sms10x10" value="N"  <% if (OUserInfo.Fitemlist(0).Fsms10x10 = "N") then %>checked<% end if %>></td>
			<td style="padding-left:2px;">아니오</td>
			</tr>
			</table>
		<% else %>
			계정없음
		<%end if %>
	</td>
	<td height="30" width="15%" bgcolor="#DDDDFF">핑거스 :</td>
	<td bgcolor="#FFFFFF" colspan="3">
		<% if (haveonlineaccount = "Y") then %>
			<table class="a" border="0" cellspacing="0" cellpadding="0">
			<tr>
				<td style="padding-bottom:2px;"><input type="radio" name="smsfinger" value="Y" <% if (OUserInfo.Fitemlist(0).Fsmsfinger = "Y") then %>checked<% end if %>></td>
				<td style="padding-left:2px;">예</td>
				<td style="padding:0 0 2px 15px;"><input type="radio"  name="smsfinger" value="N"  <% if (OUserInfo.Fitemlist(0).Fsmsfinger = "N") then %>checked<% end if %>></td>
				<td style="padding-left:2px;">아니오</td>
			</tr>
			</table>
		<% else %>
			계정없음
		<%end if %>
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">포인트 :</td>
	<td bgcolor="#FFFFFF" colspan="3" width="35%">
		<% if (haveofflineaccount = "Y") then %>
			<table class="a" border="0" cellspacing="0" cellpadding="0">
			<tr>
				<td style="padding-bottom:2px;"><input type="radio" name="offlinesms" value="Y" <% if (OOfflineUserInfo.Fitemlist(0).Fsms = "Y") then %>checked<% end if %>></td>
				<td style="padding-left:2px;">예</td>
				<td style="padding:0 0 2px 15px;"><input type="radio"  name="offlinesms" value="N"  <% if (OOfflineUserInfo.Fitemlist(0).Fsms = "N") then %>checked<% end if %>></td>
				<td style="padding-left:2px;">아니오</td>
			</tr>
			</table>
		<% else %>
			계정없음
		<%end if %>
	</td>
	<td height="30" width="15%" bgcolor="#DDDDFF"></td>
	<td bgcolor="#FFFFFF" colspan="3"></td>
</tr>
</table>
<br>
<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="center">
		<input type="button" class="button" value="저장하기" onClick="if (confirm('저장하시겠습니까?')) {document.frm.submit();}">
		<input type="button" class="button" value=" 창닫기 " onClick="self.close()">
	</td>
</tr>
</table>
<!-- 액션 끝 -->
</form>

<%
set OUserInfo = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->

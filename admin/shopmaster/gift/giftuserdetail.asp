<%@ language=vbscript %>
<% option explicit %>
<%
	Response.AddHeader "Cache-Control","no-cache"
	Response.AddHeader "Expires","0"
	Response.AddHeader "Pragma","no-cache"
%>
<%
'####################################################
' Description :  사은품대상자리스트
' History : 2019.09.25 한용민 생성
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/classes/items/itemgiftcls.asp"-->
<!-- #include virtual="/lib/classes/event/eventmanageCls.asp"-->
<!-- #include virtual="/admin/eventmanage/common/event_function.asp"-->
<!-- #include virtual="/lib/classes/event/openGiftCls.asp"-->
<%
dim gift_code, giftexistsyn, ipkumdiv, reload, cgiftuser, page, orderserial, userid, reqname
	gift_code = requestCheckVar(getNumeric(Request("gift_code")),10)
	giftexistsyn = requestCheckVar(Request("giftexistsyn"),1)
	ipkumdiv = requestCheckVar(getNumeric(Request("ipkumdiv")),10)
	reload = requestCheckVar(Request("reload"),10)
	page = requestCheckVar(getNumeric(Request("page")),10)
	orderserial = requestCheckVar(getNumeric(Request("orderserial")),11)
	userid = requestCheckVar(Request("userid"),32)
	reqname = requestCheckVar(Request("reqname"),32)

if giftexistsyn="" then giftexistsyn="N"
if page = "" then page = 1

dim clsGift, cEGroup, igScope, eCode, ieGroupCode, sBrand, igType, igR1, igR2, igkCode, igkType
dim sTitle, igkCnt, igkLimit, dSDay, dEDay, igStatus, igUsing, dRegdate, sAdminid, igkName, sgkImg
dim sgDelivery, dOpenDay, dCloseDay, sOldName, iSiteScope, sPartnerID, BCouponIdx, giftkind_linkGbn
dim giftkind_givecnt, arrlist, eregdate, GiftIsusing, GiftImage1, GiftText1, GiftImage2, GiftText2
dim GiftImage3, GiftText3, GiftInfoText, blngroup, arrGroup, i, arrsitescope, intgroup

set clsGift = new CGift
	clsGift.FGCode = gift_code

	if gift_code<>"" then
		clsGift.fnGetGiftConts
	end if

	sTitle		= clsGift.FGName
	igScope 	= clsGift.FGScope
	eCode		= clsGift.FECode
	ieGroupCode	= clsGift.FEGroupCode
	sBrand		= clsGift.FBrand
	igType		= clsGift.FGType
	igR1		= clsGift.FGRange1
	igR2 		= clsGift.FGRange2
	igkCode		= clsGift.FGKindCode
	igkType		= clsGift.FGKindType
	igkCnt		= clsGift.FGKindCnt
	igkLimit	= clsGift.FGKindlimit
	dSDay		= clsGift.FSDate
	dEDay		= clsGift.FEDate
	igStatus	= clsGift.FGStatus
	igUsing     = clsGift.FGUsing
	dRegdate	= clsGift.FRegdate
	sAdminid 	= clsGift.FAdminid
	igkName 	= clsGift.FGKindName
	sgkImg		= clsGift.FGKindImg
	sgDelivery  = clsGift.FGDelivery
	dOpenDay	= clsGift.FOpenDate
	dCloseDay	= clsGift.FCloseDate
	sOldName	= clsGift.FOldKindName
	iSiteScope	= clsGift.FSiteScope
	sPartnerID	= clsGift.FPartnerID
	BCouponIdx  = clsGift.Fbcouponidx
	giftkind_linkGbn = clsGift.Fgiftkind_linkGbn
	giftkind_givecnt = clsGift.Fgiftkind_givecnt

	If giftkind_givecnt > 0 Then ''사은품 한정제공수량
	arrlist = clsGift.fnLimitgiftCount
	End If

	eregdate = dSDay

	clsGift.FECode = eCode

	if gift_code<>"" then
		clsGift.fnGetEventGiftBox	' 이벤트 사은품 박스 정보 가져오기
	end if

	GiftIsusing = clsGift.FGiftIsusing
	GiftImage1 = clsGift.FGiftImage1
	GiftText1 = clsGift.FGiftText1
	GiftImage2 = clsGift.FGiftImage2
	GiftText2 = clsGift.FGiftText2
	GiftImage3 = clsGift.FGiftImage3
	GiftText3 = clsGift.FGiftText3
	GiftInfoText = clsGift.FGiftInfoText
set clsGift = nothing

IF eCode = 0 THEN eCode = ""
IF igkLimit = 0 THEN igkLimit = ""
IF isNull(igkLimit) THEN igkLimit = ""

IF eCode <> "" THEN	'이벤트와 연관된 사은품일 경우
	arrsitescope = fnSetCommonCodeArr("eventscope",True) '범위 코드값에 따른 명칭 가져오기
	'그룹리스트
	set cEGroup = new ClsEventGroup
		cEGroup.FECode = eCode
		arrGroup = cEGroup.fnGetEventItemGroup	' 이벤트화면설정 그룹내용가져오기
	set cEGroup = nothing
END IF
blngroup = False
IF isArray(arrGroup) THEN blngroup = True

	  '공통코드 값 배열로 한꺼번에 가져온 후 값 보여주기
	Dim  arrgiftstatus
	arrgiftstatus 	= fnSetCommonCodeArr("giftstatus",False)

  ''전체사은or 다이어리 이벤트 인지 Check -----------------
    Dim oOpenGift, iopengiftType, iopengiftName, iopengiftfrontOpen
    iopengiftType = 0
    set oOpenGift=new CopenGift
    oOpenGift.FRectEventCode = eCode
    if (eCode<>"") then
        oOpenGift.getOneOpenGift

        if (oOpenGift.FResultcount>0) then
            iopengiftType       = oOpenGift.FOneItem.FopengiftType
            iopengiftName       = oOpenGift.FOneItem.getOpengiftTypeName
            iopengiftfrontOpen  = oOpenGift.FOneItem.FfrontOpen

            igScope = iopengiftType
        end if
    end if
    set oOpenGift=Nothing
dim eFolder
eFolder=eCode

' 사은품대상자
set cgiftuser = new CGift
	cgiftuser.FPageSize = 1000
	cgiftuser.FCurrPage = page
	cgiftuser.frectgift_code = gift_code
	cgiftuser.frectgiftexistsyn = giftexistsyn
	cgiftuser.frectorderserial = orderserial
	cgiftuser.frectuserid = userid
	cgiftuser.frectreqname = reqname
	cgiftuser.frectipkumdiv = ipkumdiv

	if gift_code<>"" then
		cgiftuser.fngiftuserlist
	end if
%>
<script type="text/javascript">

function frmsubmit(page){
	frmgift.page.value=page;
	frmgift.submit();
}

function fngiftremakebefore(gift_code){
	if (gift_code==""){
		alert("사은품 코드가 없습니다.");
		return;
	}

	<% 'if C_ADMIN_AUTH then %>
		var ret = confirm("출고 이전 사은품을 재작성 합니다\n계속진행하시겠습니까?");
		if (ret) {
			frmproc.action='/admin/shopmaster/gift/giftuser_process.asp';
			frmproc.gift_code.value=gift_code;
			frmproc.mode.value='giftremakebefore';
			frmproc.submit();
		}
	<% 'else %>
		//alert("관리자만 사용가능한 매뉴 입니다.");
		//return;
	<% 'end if %>
}

function fngiftremakeafter(gift_code){
	if (gift_code==""){
		alert("사은품 코드가 없습니다.");
		return;
	}

	<% 'if C_ADMIN_AUTH then %>
		var ret = confirm("출고 이후 사은품을 서비스발송 합니다\n계속진행하시겠습니까?");
		if (ret) {
			frmproc.action='/admin/shopmaster/gift/giftuser_process.asp';
			frmproc.gift_code.value=gift_code;
			frmproc.mode.value='giftremakeafter';
			frmproc.submit();
		}
	<% 'else %>
		//alert("관리자만 사용가능한 매뉴 입니다.");
		//return;
	<% 'end if %>
}

</script>

<!-- 검색 시작 -->
<form name="frmgift" method="get" action="" style="margin:0px;">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="page" value="1">
<input type="hidden" name="reload" value="on">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		* 사은품코드 : <input type="text" name="gift_code" value="<%= gift_code %>" size="8" maxlength="9" >
		&nbsp;
		* 주문번호 : <input type="text" name="orderserial" value="<%= orderserial %>" size="10" maxlength="11" >
		&nbsp;
		* 고객아이디 : <input type="text" name="userid" value="<%= userid %>" size="10" maxlength="11" >
		&nbsp;
		* 수령인이름 : <input type="text" name="reqname" value="<%= reqname %>" size="10" maxlength="11" >
	</td>
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="frmsubmit('1');">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		* 사은품포함여부 :
		<select name="giftexistsyn">
			<option value="">전체</option>
			<option value="N" <% if giftexistsyn="N" then response.write " selected" %>>사은품미포함(누락)</option>
			<option value="Y" <% if giftexistsyn="Y" then response.write " selected" %>>사은품포함</option>
		</select>
		* 출고상태 :
		<select name="ipkumdiv">
			<option value="">전체</option>
			<option value="98" <% if ipkumdiv="98" then response.write " selected" %>>출고이전</option>
			<option value="99" <% if ipkumdiv="99" then response.write " selected" %>>출고완료</option>
		</select>
	</td>
</tr>
</table>
</form>
<!-- 검색 끝 -->
<br>
<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
		실시간 사은품 검토로 인해 부하가 심한 매뉴 입니다. 한번만 클릭하시고 기다려 주세요.
	</td>
	<td align="right">
		<% 'if C_ADMIN_AUTH then %>
			<input type="button" class="button" value="출고이전사은품재작성" onClick="fngiftremakebefore('<%= gift_code %>');">
			<input type="button" class="button" value="출고이후사은품서비스발송" onClick="fngiftremakeafter('<%= gift_code %>');">
		<% 'end if %>
	</td>
</tr>
</table>
<!-- 액션 끝 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="25">
		검색결과 : <b><%= cgiftuser.FTotalCount %></b>
		&nbsp;
		페이지 : <b><%= page %>/ <%= cgiftuser.FTotalPage %></b>
		&nbsp;&nbsp;※ 최대 10000건까지 검색 됩니다.
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>주문번호</td>
	<td>수령인</td>
	<td>아이디</td>
	<td>사은품수</td>
	<td>출고상태</td>
	<td>사은품명</td>
	<td>출고내역서</td>
	<td>사은품포함여부</td>
</tr>
<% if cgiftuser.FresultCount>0 then %>
	<% for i=0 to cgiftuser.FresultCount-1 %>
	<tr align="center" bgcolor="#FFFFFF">
		<td><%= cgiftuser.FItemList(i).forderserial %></td>
		<td><%= cgiftuser.FItemList(i).freqname %></td>
		<td><%= cgiftuser.FItemList(i).fuserid %></td>
		<td><%= cgiftuser.FItemList(i).fgiftkind_cnt %></td>
		<td><%= cgiftuser.FItemList(i).fipkumdivname %></td>
		<td><%= cgiftuser.FItemList(i).fgift_name %></td>
		<td><%= cgiftuser.FItemList(i).fEventConditionStr %></td>
		<td>
			<% if cgiftuser.FItemList(i).fgiftexistsyn="Y" then %>
				<strong>사은품포함</strong>
			<% else %>
				<% if cgiftuser.FItemList(i).fgiftserviceyn="Y" then %>
					<strong>서비스발송</strong>
				<% else %>
					<strong><font color="red">사은품미포함(누락)</font></strong>
				<% end if %>
			<% end if %>
	</tr>
	<% next %>

	<tr height="25" bgcolor="FFFFFF">
		<td colspan="16" align="center">
			<% if cgiftuser.HasPreScroll then %>
				<span class="list_link"><a href="javascript:frmsubmit(<%= cgiftuser.StartScrollPage-1 %>)">[pre]</a></span>
			<% else %>
			[pre]
			<% end if %>
			<% for i = 0 + cgiftuser.StartScrollPage to cgiftuser.StartScrollPage + cgiftuser.FScrollCount - 1 %>
				<% if (i > cgiftuser.FTotalpage) then Exit for %>
				<% if CStr(i) = CStr(cgiftuser.FCurrPage) then %>
				<span class="page_link"><font color="red"><b><%= i %></b></font></span>
				<% else %>
				<a href="javascript:frmsubmit(<%= i %>)" class="list_link"><font color="#000000"><%= i %></font></a>
				<% end if %>
			<% next %>
			<% if cgiftuser.HasNextScroll then %>
				<span class="list_link"><a href="javascript:frmsubmit(<%= i %>)">[next]</a></span>
			<% else %>
			[next]
			<% end if %>
		</td>
	</tr>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="25" align="center" class="page_link">
			<% if gift_code="" then %>
				<font color="red">사은품 코드를 입력해주셔야 검색이 됩니다.</font>
			<% else %>
				[검색결과가 없습니다.]
			<% end if %>
		</td>
	</tr>
<% end if %>

</table>
<form name="frmproc" method="post" action="" style="margin:0px;">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="gift_code" value="<%= gift_code %>">
<input type="hidden" name="mode" value="">
</form>

<%
set cgiftuser = nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->

<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : cs센터
' History : 2009.04.17 이상구 생성
'			2016.06.30 한용민 수정
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/checkAllowIPWithLog.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/cscenter/cs_boardusercls.asp" -->
<%
dim userid, i

dim occscenterboarduser
set occscenterboarduser = new CCSCenterBoardUser
	occscenterboarduser.FPageSize = 50
	occscenterboarduser.FCurrPage = 1
	occscenterboarduser.GetCSCenterBoardUserList

'// 마스터 이상:2 및 시스템팀:7
if Not ((session("ssAdminLsn") <= 2) or (session("ssAdminPsn") = 7) or C_CSPowerUser or C_ADMIN_AUTH or C_CSpermanentUser) then
	response.write "<br><br>권한이 없습니다."
	response.end
end if

dim IsSystemPsn	: IsSystemPsn = False
if (session("ssAdminPsn") = 7) then
	IsSystemPsn = True
end if

%>
<script type="text/javascript">

function ModifyIppbxInfo(frm){
	if ((frm.userid.value == "") && (frm.useyn.value == "Y")) {
		alert("아이디를 지정하세요.\n\n또는 사용안함으로 수정하세요.");
		return;
	}

	if (confirm("수정하시겠습니까?") == true) {
		frm.submit();
	}
}

function doVVipReOrganize(frm){
	if (confirm("VVIP상담 담당자 수동분배 하시겠습니까?") == true) {
		frm.mode.value = "reorganizevvipone2one";
		frm.submit();
	}
}

function doVipReOrganize(frm){
	if (confirm("VIP상담 담당자 수동분배 하시겠습니까?") == true) {
		frm.mode.value = "reorganizevipone2one";
		frm.submit();
	}
}

function doVipReOrganizeNoCharge(frm){
	if (confirm("미분배 VIP상담 담당자 수동분배 하시겠습니까?") == true) {
		frm.mode.value = "reorganizevipone2onenocharge";
		frm.submit();
	}
}

function doReOrganize(frm){
	if (confirm("일반상담 담당자 수동분배 하시겠습니까?") == true) {
		frm.submit();
	}
}

function doReOrganizeNoCharge(frm) {
	if (confirm("미분배 일반상담 담당자 수동분배 하시겠습니까?") == true) {
		frm.mode.value = "reorganizeone2onenocharge";
		frm.submit();
	}
}

function doReOrganizeMichulgoNoCharge(frm){
	if (confirm("미분배 브랜드 담당자 수동분배 하시겠습니까?") == true) {
		frm.mode.value = "reorganizemichulgonocharge";
		frm.submit();
	}
}

function doReOrganizeNotReturnNoCharge(frm){
	if (confirm("미분배 브랜드 담당자 수동분배 하시겠습니까?") == true) {
		frm.mode.value = "reorganizenotreturnnocharge";
		frm.submit();
	}
}

function doReOrganizeMichulgoAll(frm){
	if (confirm("전체 담당자 수동분배 하시겠습니까?") == true) {
		frm.mode.value = "reorganizemichulgoall";
		frm.submit();
	}
}

function doReOrganizeNotReturnAll(frm){
	if (confirm("전체 담당자 수동분배 하시겠습니까?") == true) {
		frm.mode.value = "reorganizenotreturnall";
		frm.submit();
	}
}

function doReOrganizeMichulgoAvgAll(frm){
	if (confirm("담당자 브랜드 덜어주기 수동분배 하시겠습니까?") == true) {
		frm.mode.value = "reorganizemichulgoavg";
		frm.submit();
	}
}

function doReOrganizeStockout(frm){
	if (confirm("품절출고불가 전체 재분배 하시겠습니까?") == true) {
		frm.mode.value = "reorganizestockoutall";
		frm.submit();
	}
}

</script>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		<!--아이디 : <input type="text" class="text" name="userid" value="<%= userid %>">-->
	</td>
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
      	<input type="button" class="button_s" value="검색" onclick="document.frm.submit()">
	</td>
</tr>
</form>
</table>

<br>

* 휴가상태는 휴가달력(휴가 <font color=red>신청상태 포함</font>)에서 가져옵니다.<br>
* 휴가상태가 휴가중 이거나, 사용안함 설정하면 담당자 자동분배에서 제외됩니다.<br>
* 1:1 상담 담당자 분배는 <font color=red>고객 글 작성시</font> 자동분배됩니다.<br><br>

* <font color=red>담당자 수동분배</font>를 하여 수기로 분배 할 수 있습니다.

<br>

<b>* 1:1 상담 게시판</b>
<input type=button class=button value="VVIP 상담 전체 재분배" onClick="doVVipReOrganize(frmAction)">
&nbsp;
<input type=button class=button value="VIP 상담 전체 재분배" onClick="doVipReOrganize(frmAction)">
&nbsp;
&nbsp;
&nbsp;
&nbsp;
&nbsp;
&nbsp;
&nbsp;
&nbsp;
&nbsp;
<input type=button class=button value="VIP 상담 미지정 재분배" onClick="doVipReOrganizeNoCharge(frmAction)">
&nbsp;
&nbsp;
&nbsp;
&nbsp;
&nbsp;
&nbsp;
&nbsp;
&nbsp;
&nbsp;
<input type=button class=button value="일반상담 미지정 재분배" onClick="doReOrganizeNoCharge(frmAction)">
&nbsp;
&nbsp;
&nbsp;
&nbsp;
&nbsp;
&nbsp;
&nbsp;
&nbsp;
&nbsp;
<input type=button class=button value="일반상담 전체 재분배" onClick="doReOrganize(frmAction)">
<% if IsSystemPsn then %>
	<input type=button class=button value="전체 재분배[시스템팀]" onClick="doReOrganize(frmAction)">
<% end if %>
<br><br>

<b>* D+3 미발송건(업배)</b> <input type=button class=button value="미지정건 분배" onClick="doReOrganizeMichulgoNoCharge(frmAction)">
<input type=button class=button value="전체 재분배" onClick="doReOrganizeMichulgoAll(frmAction)">
<input type=button class=button value="지정된 브랜드 덜어주기" onClick="doReOrganizeMichulgoAvgAll(frmAction)">

<br><br>

<b>* 품절취소요청건(입점몰제외, 고객 안내이전)</b>
<input type=button class=button value="전체 재분배<% if (C_CSPowerUser or C_ADMIN_AUTH) then %>[관리자권한]<% end if %>" onClick="doReOrganizeStockout(frmAction)" <% if Not(C_CSPowerUser or C_ADMIN_AUTH) then %>disabled<% end if %> >

<br><br>

<b>* D+3, D+7 반품 미처리건(업배)</b> <input type=button class=button value="미지정건 분배" onClick="doReOrganizeNotReturnNoCharge(frmAction)">
<input type=button class=button value="전체 재분배" onClick="doReOrganizeNotReturnAll(frmAction)">

<br>

<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmAction" method="post" action="boarduser_process.asp">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="mode" value="reorganizeone2one">
</form>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="25">
	<td width="60">순서</td>
    <td width="150">아이디</td>
    <td width="80">휴가상태</td>
    <td width="90"><font color="red">VVIP</font> 1:1상담</td>
    <td width="90">VIP 1:1상담</td>
	<td width="90">1:1상담</td>
    <td width="90">미출고</td>
    <td width="90">품절취소</td>
    <td width="90">반품</td>
    <td width="90">사용</td>
    <td width="150">수정일</td>
    <td>비고</td>
</tr>
<% if occscenterboarduser.FTotalCount > 0 then %>
	<% for i = 0 to (occscenterboarduser.FResultCount - 1) %>

	<% if (occscenterboarduser.FItemList(i).Fuseyn = "N") then %>
		<tr align="center" bgcolor="#DDDDDD" height="25">
	<% else %>
		<tr align="center" bgcolor="#FFFFFF" height="25">
	<% end if %>

		<form name="frm<%= i %>" method="post" action="/cscenter/board/boarduser_process.asp">
		<input type="hidden" class="text" name="menupos" value="<%= menupos %>">
		<input type="hidden" name="mode" value="modify">
		<input type="hidden" class="text" name="indexno" value="<%= occscenterboarduser.FItemList(i).Findexno %>">
	    <td><%= occscenterboarduser.FItemList(i).Findexno %></td>
	    <td><input type="text" class="text" name="userid" value="<%= occscenterboarduser.FItemList(i).Fuserid %>" size="16"></td>
	    <td>
	    	<% if (occscenterboarduser.FItemList(i).Fuserid <> "") then %>
	        	<% if (occscenterboarduser.FItemList(i).Fvacationyn = "Y") then %>
	        		<font color=red>휴가중</font>
	        	<% else %>
	        		근무중
	        	<% end if %>
	        <% end if %>
	    </td>
	    <td bgcolor="#ABF200">
			<select name="vvipone2oneyn" class="select">
				<option value="Y" <% if (occscenterboarduser.FItemList(i).fvvipone2oneyn = "Y") then %>selected<% end if %>>분배함
				<option value="T" <% if (occscenterboarduser.FItemList(i).fvvipone2oneyn = "T") then %>selected<% end if %>>분배정지
				<option value="N" <% if (occscenterboarduser.FItemList(i).fvvipone2oneyn = "N") then %>selected<% end if %>>분배안함
			</select>
	    </td>
	    <td>
			<select name="vipone2oneyn" class="select">
				<option value="Y" <% if (occscenterboarduser.FItemList(i).Fvipone2oneyn = "Y") then %>selected<% end if %>>분배함
				<option value="T" <% if (occscenterboarduser.FItemList(i).Fvipone2oneyn = "T") then %>selected<% end if %>>분배정지
				<option value="N" <% if (occscenterboarduser.FItemList(i).Fvipone2oneyn = "N") then %>selected<% end if %>>분배안함
			</select>
	    </td>
	    <td>
			<select name="one2oneyn" class="select">
				<option value="Y" <% if (occscenterboarduser.FItemList(i).Fone2oneyn = "Y") then %>selected<% end if %>>분배함
				<option value="T" <% if (occscenterboarduser.FItemList(i).Fone2oneyn = "T") then %>selected<% end if %>>분배정지
				<option value="N" <% if (occscenterboarduser.FItemList(i).Fone2oneyn = "N") then %>selected<% end if %>>분배안함
			</select>
	    </td>
	    <td>
			<select name="michulgoyn" class="select">
				<option value="Y" <% if (occscenterboarduser.FItemList(i).Fmichulgoyn = "Y") then %>selected<% end if %>>분배함
				<option value="T" <% if (occscenterboarduser.FItemList(i).Fmichulgoyn = "T") then %>selected<% end if %>>분배정지
				<option value="N" <% if (occscenterboarduser.FItemList(i).Fmichulgoyn = "N") then %>selected<% end if %>>분배안함
			</select>
	    </td>
	    <td>
			<select name="stockoutyn" class="select">
				<option value="Y" <% if (occscenterboarduser.FItemList(i).Fstockoutyn = "Y") then %>selected<% end if %>>분배함
				<option value="T" <% if (occscenterboarduser.FItemList(i).Fstockoutyn = "T") then %>selected<% end if %>>분배정지
				<option value="N" <% if (occscenterboarduser.FItemList(i).Fstockoutyn = "N") then %>selected<% end if %>>분배안함
			</select>
	    </td>
	    <td>
			<select name="returnyn" class="select">
				<option value="Y" <% if (occscenterboarduser.FItemList(i).Freturnyn = "Y") then %>selected<% end if %>>분배함
				<option value="T" <% if (occscenterboarduser.FItemList(i).Freturnyn = "T") then %>selected<% end if %>>분배정지
				<option value="N" <% if (occscenterboarduser.FItemList(i).Freturnyn = "N") then %>selected<% end if %>>분배안함
			</select>
	    </td>
	    <td>
			<select name="useyn" class="select">
				<option value="Y" <% if (occscenterboarduser.FItemList(i).Fuseyn = "Y") then %>selected<% end if %>>사용함
				<option value="N" <% if (occscenterboarduser.FItemList(i).Fuseyn = "N") then %>selected<% end if %>>사용안함
			</select>
	    </td>
	    <td><%= occscenterboarduser.FItemList(i).Flastupdate %></td>
	    <td>
	    	<input type="button" class="button" value="수정" onClick="ModifyIppbxInfo(frm<%= i %>)">
	    </td>
	    </form>
	</tr>
	<% next %>
<% else %>
	<tr bgcolor="#FFFFFF" align="center">
	    <td height="25" colspan="13">검색결과가 없습니다.</td>
	</tr>
<% end if %>

</table>

<%
set occscenterboarduser = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->

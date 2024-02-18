<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : cs센터 고객조회
' History : 2009.04.17 이상구 생성
'           2023.10.30 한용민 수정(휴면계정정보표기. 휴면계정->일반계정 전환 로직 생성.)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/checkAllowIPWithLog.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/util/datelib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/member/customercls.asp"-->
<!-- #include virtual="/lib/classes/member/offlinecustomercls.asp"-->
<%
dim isonline, mode, searchText, ishold, currpage, i, j, buf, isdisphold, snsGubun, snsGubunList, OUserSnsInfo
dim OUserInfoList
	mode 		= requestCheckvar(trim(request("mode")), 32)
	searchText 	= requestCheckvar(trim(request("searchText")),128)
	currpage 	= requestCheckvar(trim(getNumeric(request("currpage"))),8)
	isonline 	= requestCheckvar(trim(request("isonline")),1)
	ishold 		= requestCheckvar(trim(request("ishold")),1)

if (mode = "") then
	mode = "id"
end if

if (currpage = "") then
	currpage = 1
end if

if (isonline = "") then
	isonline = "Y"
end if

if (isonline = "Y") then
	set OUserInfoList = new CUserInfo
else
	set OUserInfoList = new COfflineUserInfo
end if

' ISMS 심사로 인해 휴면계정, 특정사람만 보이게(cs팀이고 개인정보취급자 이거나 한용민,허진원)	' 2020.09.21 한용민
if (C_CSUser and C_CriticInfoUserLV1) or C_privacyadminuser then
	isdisphold = true
else
	isdisphold = false
	ishold = ""
end if

OUserInfoList.FPageSize = 50
OUserInfoList.FCurrPage = currpage
OUserInfoList.FRectMode = mode
OUserInfoList.FRectHoldUser = ishold

select case mode
	case "id"
		OUserInfoList.FRectUserID = searchText
	case "partid"
		OUserInfoList.FRectUserID = searchText
	case "name"
		OUserInfoList.FRectUserName = searchText
	case "cell"
		OUserInfoList.FRectUserCell = searchText
	case "mail"
		OUserInfoList.FRectUserMail = searchText
	case else
		''
end select

if (searchText = "") then
	OUserInfoList.FresultCount = 0
else
	OUserInfoList.GetUserList
end if

%>
<script type="text/javascript">

function SubmitForm(){
	if (frm.searchText.value!=""){
		if (frm.mode.value=="cell"){
			if (instr(frm.searchText.value,"@")>0){
				alert("휴대폰번호를 정확하게 입력해 주세요.");
				return;
			}
		}
		if (frm.mode.value=="mail"){
			if (instr(frm.searchText.value,"@")<1){
				alert("이메일주소를 정확하게 입력해 주세요.");
				return;
			}
		}
	}
	document.frm.submit();
}

function openWindowMemberDetail(userid, userseq){
	var pop = window.open("/cscenter/member/popcustomerview.asp?userid=" + userid + "&userseq=" + userseq,"WindowMemberDetail","width=1400 height=800 scrollbars=yes resizable=yes");
	pop.focus();
}

function ResetUserPass(frm, userid) {
	if (confirm("\n\n주의!!!!\n\n임시 비밀번호를 생성합니다.\n\n임시비밀번호는 자동으로 발송되지 않으며 CS메모에만 기록됩니다.\n(별도 고객안내 필요)\n\n진행하시겠습니까?") == true) {
		frm.mode.value = "resetUserPass";
		frm.userid.value = userid;
		frm.target="";
		frm.submit();
	}
}

<% ' 기능은 있으나, isms 결함조치로 막음 %>
//function popDelonUser(userid, userseq){
<% '	var popDel = window.open("/cscenter/member/popcustomerdel.asp?userid=" + userid + "&userseq=" + userseq,"DelDetail","width=1400 height=800 scrollbars=yes resizable=yes"); %>
//	popDel.focus();
//}

function popChangeOnHoldUser(userid, userseq){
	if (confirm('온라인 휴면 고객을 일반회원으로 전환 합니다.\n진행하시겠습니까?') == true) {
		frmAct.userid.value = userid;
		frmAct.mode.value = "ChangeOnHoldUser";
		frmAct.target="view";
		frmAct.action = "/cscenter/member/domodifyuserinfo.asp";
		frmAct.submit();
	}
}

</script>

<!-- 검색 시작 -->
<form name="frm" method="get" action="" style="margin:0px;">
<input type="hidden" name="currpage" value="1">
<input type="hidden" name="menupos" value="<%= request("menupos") %>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		<select class="select" name="mode">
			<option value="id" <%=chkIIF(mode="id","selected","")%>>아이디</option>
			<!-- 무지 느림
			<option value="partid" <%=chkIIF(mode="partid","selected","")%>>아이디(일부분)</option>
			-->
			<option value="name" <%=chkIIF(mode="name","selected","")%>>이름</option>
			<option value="cell" <%=chkIIF(mode="cell","selected","")%>>핸드폰</option>
			<option value="mail" <%=chkIIF(mode="mail","selected","")%>>이메일</option>
		</select>
		&nbsp;
		<input type="text" class="text" name="searchText" value="<%= searchText %>" size="32" onKeyPress="if (event.keyCode == 13) SubmitForm();">
		<% if isdisphold then %>
			&nbsp;
			<input type="checkbox" name="ishold" value="Y" <% if (ishold = "Y") then %>checked<% end if %> > 휴면계정검색
		<% end if %>
	</td>

	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="javascript:SubmitForm();">
	</td>
</tr>
</table>
</form>

<br>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="16">
		검색결과 : <b>총 <%= OUserInfoList.FTotalCount %> 건</b>
		&nbsp;
		페이지 : <b><%= currpage %> / <%= OUserInfoList.FTotalPage %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="50" align="center">구분</td>
	<td width="60" align="center">등급</td>
	<td width="80" align="center">아이디</td>
	<td width="80">고객명</td>
	<td width="90" align="center">회원가입일</td>
	<td>이메일</td>
	<td width="100">전화번호</td>
	<td width="100">핸드폰번호</td>
	<td width="30">실명<br>인증</td>
	<td width="50">휴면계정</td>
	<td width="80">회원가입방식</td>
	<td width="150">비고</td>
</tr>

<% if OUserInfoList.FresultCount < 1 then %>
	<tr align="center" bgcolor="#FFFFFF">
		<td colspan="13" align="center">[검색결과가 없습니다.]</td>
	</tr>
<% else %>
	<% for i = 0 to OUserInfoList.FresultCount - 1 %>
	<tr align="center" bgcolor="#FFFFFF">
		<td><% if (isonline = "Y") then response.write "온라인" else response.write "오프라인" end if %></td>
		<td>
			<% if (isonline = "Y") then %>
				<font color="<%= getUserLevelColorByDate(OUserInfoList.FItemList(i).fUserLevel, date()) %>">
				<%= getUserLevelStrByDate(OUserInfoList.FItemList(i).fUserLevel, date()) %></font>
			<% end if %>
		</td>
		<td><%= OUserInfoList.FItemList(i).FUserID %></td>
		<td><%= OUserInfoList.FItemList(i).FUserName %></td>
		<td><%= Left(OUserInfoList.FItemList(i).Fregdate,10) %></td>
		<td>
			<%
		  if OUserInfoList.FItemList(i).FUsermail <> "" and not(isnull(OUserInfoList.FItemList(i).FUsermail)) then
			if (Len(OUserInfoList.FItemList(i).FUsermail) > 0) then
				buf = Split(OUserInfoList.FItemList(i).FUsermail, "@")
				if (UBound(buf) < 1) then
					response.write OUserInfoList.FItemList(i).FUsermail
				else
					if (Len(buf(0)) > 3) then
						response.write Left(buf(0), (Len(buf(0)) - 3)) & "***" & "@" & buf(1)
					else
						response.write buf(0) & "@" & buf(1)
					end if
				end if
			end if
		end if
		%>
		</td>
		<td>
			<%
			if OUserInfoList.FItemList(i).Fuserphone <> "" and not(isnull(OUserInfoList.FItemList(i).Fuserphone)) then
				if (Len(OUserInfoList.FItemList(i).Fuserphone) > 3) then
					response.write AstarPhoneNumber(OUserInfoList.FItemList(i).Fuserphone)
				else
					response.write OUserInfoList.FItemList(i).Fuserphone
				end if
			end if
			%>
		</td>
		<td>
			<%
			if OUserInfoList.FItemList(i).Fusercell <> "" and not(isnull(OUserInfoList.FItemList(i).Fusercell)) then
				if (Len(OUserInfoList.FItemList(i).Fusercell) > 3) and (ishold <> "Y") then
					if (Left(Now, 10) >= "2014-04-21") and (Left(Now, 10) < "2014-04-22") then
						'// TODO : 특정 기간만 핸드폰번호 전체 표시
						response.write OUserInfoList.FItemList(i).Fusercell
					else
						response.write AstarPhoneNumber(OUserInfoList.FItemList(i).Fusercell)
					end if
				else
					'if C_CriticInfoUserLV1 then
					'	response.write OUserInfoList.FItemList(i).Fusercell
					'else
						response.write AstarPhoneNumber(OUserInfoList.FItemList(i).Fusercell)
					'end if
				end if
			end if
			%>
		</td>
		<td>
			<% if (isonline = "Y") then %>
				<%= OUserInfoList.FItemList(i).Frealnamecheck %>
			<% end if %>
		</td>
		<td>
			<% if OUserInfoList.FItemList(i).fHoldUseryn="Y" then %>
				휴면
			<% else %>
				일반회원
			<% end if %>
		</td>
		<td>
			<%
			if (OUserInfoList.FItemList(i).Fuserdiv = "01") then
				response.write "일반회원"
			elseif (OUserInfoList.Fitemlist(i).Fuserdiv = "05") then
				set OUserSnsInfo = new CUserInfo
					OUserSnsInfo.FRectUserID = OUserInfoList.FItemList(i).FUserID
					snsGubunList = OUserSnsInfo.GetSNSUserJoinPathList
				set OUserSnsInfo = nothing
				if isArray(snsGubunList) then
					for j=0 to UBound(snsGubunList,2)
						snsGubun = snsGubun & chkIIF(snsGubun<>""," / ","") & GetSNSJoinTypeName(snsGubunList(0,j))
					Next
				end if
				response.write "SNS가입회원<br>(" & snsGubun & ")"
			elseif (OUserInfoList.FItemList(i).Fuserdiv = "96") then
				response.write "차단 기타 회원 (정지회원)"
			end if
			%>
		</td>
		<td>
			<input type="button" class="button" value="변경" onclick="openWindowMemberDetail('<%= OUserInfoList.FItemList(i).FUserID %>', '<%= OUserInfoList.FItemList(i).FUserSeq %>')" <% if isonline="Y" and ishold="Y" then %>disabled<% end if %> >
			<!--<input type="button" class="button" value="탈퇴처리" onClick="popDelonUser('<%'= OUserInfoList.FItemList(i).FUserID %>', '<%'= OUserInfoList.FItemList(i).FUserSeq %>');">-->
			<% if (isonline = "Y") and (ishold = "Y") then %>
				<%
				' 일반회원 이고 휴면계정이 아닌거
				if OUserInfoList.FItemList(i).fHoldUseryn="N" then
					if OUserInfoList.FItemList(i).Fuserdiv = "01" then
				%>
						&nbsp;
						<input type="button" class="button" value="임시비밀번호 생성" onClick="ResetUserPass(frmAct, '<%= OUserInfoList.FItemList(i).FUserID %>')">
				<%
					end if

				' 휴면회원 일때만
				elseif OUserInfoList.FItemList(i).fHoldUseryn="Y" then
				%>
					<input type="button" class="button" value="휴면->일반 회원전환" onClick="popChangeOnHoldUser('<%= OUserInfoList.FItemList(i).FUserID %>', '<%= OUserInfoList.FItemList(i).FUserSeq %>');">
				<% end if %>
			<% end if %>&nbsp;
		</td>
	</tr>
	<% next %>
<% end if %>
	<tr bgcolor="#FFFFFF">
		<td colspan="22" align="center">
    		<% if OUserInfoList.HasPreScroll then %>
    			<a href="?currpage=<%= OUserInfoList.StartScrollPage-1 %>&mode=<% =mode %>&menupos=<%= menupos %>&searchText=<%= server.UrlEncode(searchText) %>&isonline=<%= isonline %>">[pre]</a>
    		<% else %>
    			[pre]
    		<% end if %>

    		<% for i = (0 + OUserInfoList.StartScrollPage) to (OUserInfoList.FScrollCount + OUserInfoList.StartScrollPage - 1) %>
    			<% if i>OUserInfoList.FTotalpage then Exit for %>
    			<% if CStr(currpage)=CStr(i) then %>
    			<font color="red">[<%= i %>]</font>
    			<% else %>
    			<a href="?currpage=<%= i %>&mode=<% =mode %>&menupos=<%= menupos %>&searchText=<%= server.UrlEncode(searchText) %>&isonline=<%= isonline %>">[<%= i %>]</a>
    			<% end if %>
    		<% next %>

    		<% if OUserInfoList.HasNextScroll then %>
    			<a href="?currpage=<%= i %>&mode=<% =mode %>&menupos=<%= menupos %>&searchText=<%= server.UrlEncode(searchText) %>&isonline=<%= isonline %>">[next]</a>
    		<% else %>
    			[next]
    		<% end if %>
		</td>
	</tr>
</table>
<form name="frmAct" method="post" action="/cscenter/member/domodifyuserinfo.asp" onsubmit="return false;" style="margin:0px;">
	<input type="hidden" name="mode" value="">
	<input type="hidden" name="userid" value="">
</form>
<% IF application("Svr_Info")="Dev" THEN %>
	<iframe id="view" name="view" src="" width="100%" height="300"></iframe>
<% else %>
	<iframe id="view" name="view" src="" width="100%" height="0"></iframe>
<% end if %>
<%
set OUserInfoList = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->

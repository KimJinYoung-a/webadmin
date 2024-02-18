<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 고객센터 [안내문구] 기본 카테고리
' History : 이상구 생성
'			2021.09.10 한용민 수정(이문재이사님요청 자사몰 필드추가, 소스표준화, 보안강화)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/cscenter/cs_replycls.asp"-->
<%
' 위탁업체 팀장 빼고는 위탁업체 일반직원은 권한 없음.
if C_CSOutsourcingUser then
	if not(C_CSOutsourcingPowerUser) then
		response.write "권한이 없습니다."
		response.end : dbget.close()
	end if
end if

dim gubunCode : gubunCode = "0001"
dim masterUseYN, detailUseYN, masterIdx, i, page, sitename
	menupos = requestcheckvar(getNumeric(request("menupos")),10)
	page = requestcheckvar(getNumeric(request("page")),10)
	masterUseYN = requestcheckvar(request("masterUseYN"),1)
	detailUseYN = requestcheckvar(request("detailUseYN"),1)
	masteridx = requestcheckvar(getNumeric(request("masterIdx")),10)
	sitename = requestcheckvar(request("sitename"),32)

if (page = "") then
	page = 1
	masterUseYN = "Y"
	detailUseYN = "Y"
end if

dim oCReply
Set oCReply = new CReply
	oCReply.FPageSize = 50
	oCReply.FCurrPage = page
	oCReply.FRectMasterIDX = masterIdx
	oCReply.FRectMasterUseYN = masterUseYN
	oCReply.FRectDetailUseYN = detailUseYN
	oCReply.FRectGubunCode = gubunCode
	oCReply.FRectsitename = sitename
	oCReply.GetReplyDetailList()

%>

<script type="text/javascript">

function fnRegReplyDetail() {
	document.location.href = "/cscenter/board/cs_replydetail_view.asp?menupos=<%= menupos %>&masterIdx=<%= masterIdx %>&gubunCode=<%= gubunCode %>&masterUseYN=<%= masterUseYN %>";
}

function fnModiReplyDetail(idx) {
	document.location.href = "/cscenter/board/cs_replydetail_view.asp?menupos=<%= menupos %>&idx=" + idx;
}

function jsGotoPage(page) {
	var frm = document.frm;
	frm.page.value = page;
	frm.submit();
}

</script>

<!-- 검색 시작 -->
<form name="frm" method="get" action="" style="margin:0px;">
<input type="hidden" name="page" value="1">
<input type="hidden" name="menupos" value="<%= menupos %>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="1" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		* 구분 : <% Drawreplysitename "sitename", sitename, "" %>
		&nbsp;
		* 기본 카테고리:
		<% Call drawSelectBoxReplyMaster("masterIdx", masterIdx, gubunCode, masterUseYN) %>
		&nbsp;
		* 기본 카테고리 사용구분:
		<select class="select" name="masterUseYN">
			<option></option>
			<option value="Y" <% if (masterUseYN = "Y") then %>selected<% end if %> >사용함</option>
			<option value="N" <% if (masterUseYN = "N") then %>selected<% end if %> >사용안함</option>
		</select>
		&nbsp;
		* 상세 카테고리 사용구분:
		<select class="select" name="detailUseYN">
			<option></option>
			<option value="Y" <% if (detailUseYN = "Y") then %>selected<% end if %> >사용함</option>
			<option value="N" <% if (detailUseYN = "N") then %>selected<% end if %> >사용안함</option>
		</select>
	</td>
	<td rowspan="1" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="frm.submit();">
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
		</td>
		<td align="right">
			<input type="button" onclick="fnRegReplyDetail();" value="신규등록" class="button">
		</td>
	</tr>
</table>
<!-- 액션 끝 -->

<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15">
			검색결과 : <b><%= oCReply.FTotalCount %></b>
			&nbsp;
			페이지 : <b><%= page %>/ <%= oCReply.FTotalPage %></b>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="30">
		<td width="60">IDX</td>
		<td width="40">구분</td>
		<td align="left" width="150">기본 카테고리명</td>
		<td align="left" width="250">상세 카테고리명</td>
		<td align="left" width="350">안내문구</td>
		<td width="60">표시순서</td>
		<td width="30">사용</td>
		<td width="150">최종수정일</td>
		<td width="100">비고</td>
    </tr>
	<% if oCReply.FresultCount>0 then %>
	<% for i=0 to oCReply.FresultCount-1 %>
	<form action="" name="frmBuyPrc<%=i%>" method="get">

    <% if oCReply.FItemList(i).FuseYN = "Y" then %>
    <tr align="center" bgcolor="#FFFFFF" onmouseover=this.style.background="<%= adminColor("tabletop") %>"; onmouseout=this.style.background='#FFFFFF';>
    <% else %>
    <tr align="center" bgcolor="#CCCCCC">
	<% end if %>
		<td height="25">
			<a href="javascript:fnModiReplyDetail(<%= oCReply.FItemList(i).Fidx %>)"><%= oCReply.FItemList(i).Fidx %></a>
		</td>
		<td><%= replysitename(oCReply.FItemList(i).fsitename) %></td>
		<td align="left">
			<a href="javascript:fnModiReplyDetail(<%= oCReply.FItemList(i).Fidx %>)"><%= oCReply.FItemList(i).Ftitle %></a>
		</td>
		<td align="left">
			<a href="javascript:fnModiReplyDetail(<%= oCReply.FItemList(i).Fidx %>)"><%= oCReply.FItemList(i).Fsubtitle %></a>
		</td>
		<td align="left">
			<a href="javascript:fnModiReplyDetail(<%= oCReply.FItemList(i).Fidx %>)"><%= Left(oCReply.FItemList(i).Fcontents, 30) %>...</a>
		</td>
		<td>
			<%= oCReply.FItemList(i).FdispOrderNo %>
		</td>
		<td>
			<%= oCReply.FItemList(i).FuseYN %>
		</td>
		<td>
			<%= oCReply.FItemList(i).Flastupdate %>
		</td>
		<td></td>
    </tr>
	</form>
	<% next %>
	<% else %>
		<tr bgcolor="#FFFFFF">
			<td colspan="10" align="center" class="page_link">[검색결과가 없습니다.]</td>
		</tr>
	<% end if %>
    <tr height="25" bgcolor="FFFFFF">
		<td colspan="15" align="center">
	       	<% if oCReply.HasPreScroll then %>
				<span class="list_link"><a href="javascript:jsGotoPage(<%= i %>)">[pre]</a></span>
			<% else %>
			[pre]
			<% end if %>
			<% for i = 0 + oCReply.StartScrollPage to oCReply.StartScrollPage + oCReply.FScrollCount - 1 %>
				<% if (i > oCReply.FTotalpage) then Exit for %>
				<% if CStr(i) = CStr(oCReply.FCurrPage) then %>
				<span class="page_link"><font color="red"><b><%= i %></b></font></span>
				<% else %>
				<a href="javascript:jsGotoPage(<%= i %>)" class="list_link"><font color="#000000"><%= i %></font></a>
				<% end if %>
			<% next %>
			<% if oCReply.HasNextScroll then %>
				<span class="list_link"><a href="javascript:jsGotoPage(<%= i %>)">[next]</a></span>
			<% else %>
			[next]
			<% end if %>
		</td>
	</tr>
</table>

<%
set oCReply = nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->

<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/classes/organizer/organizer_cls.asp"-->
<%
'#######################################################
'	History	:  2008.10.23 한용민 생성
'	Description : 오거나이저
'#######################################################
%>
<%

dim CateCode , yearUse , isusing ,sBrand , arrItemid
CateCode = request("cate")
yearUse = "2009"
isusing = request("isusingbox")
sBrand = request("ebrand")
arrItemid = request("aitem")
dim page , i
	page = requestCheckVar(request("page"),5)
	if page = "" then page = 1

dim oip
set oip = new organizerCls
	oip.FPageSize = 50
	oip.FCurrPage = page	
	oip.frectcate = CateCode	
	oip.frectisusing = isusing		
	oip.FrectMakerid = sBrand
	oip.FRectArrItemid = arrItemid
	oip.getorganizerList

%>
<script language="javascript">

//신규 등록 팝업
function popRegNew(){
	var popRegNew = window.open('/admin/organizer/organizerReg.asp','popRegNew','width=600,height=600,status=yes')
	popRegNew.focus();
}

//상품후기 팝업
function popRegeval(itemid){
	var popRegeval = window.open('/admin/organizer/eval_list.asp?itemid='+itemid,'popRegeval','width=1024,height=768,scrollbars=yes,resizable=yes')
	popRegeval.focus();
}

//수정 팝업
function popRegModi(idx){
	var popRegModi = window.open('/admin/organizer/organizerReg.asp?mode=edit&id='+ idx,'popRegModi','width=600,height=600')
	popRegModi.focus();
}

function contents_option(){
	var contents_option = window.open('/admin/organizer/imagemake/imagemake_list.asp','contents_option','width=1024,height=768,scrollbars=yes,resizable=yes');
	contents_option.focus();
}

function keyword_option(){
	var keyword_option = window.open('/admin/organizer/option/keyword_option.asp','keyword_option','width=1024,height=768,scrollbars=yes,resizable=yes');
	keyword_option.focus();
}

function detail_view(DiaryID){
	var detail_view = window.open('/admin/organizer/option/detail_option.asp?DiaryID='+DiaryID,'detail_view','width=1024,height=768,scrollbars=yes,resizable=yes');
	detail_view.focus();
}

function edit(id){
	document.location.href="/admin/organizer/organizerReg.asp?mode=edit&id="+id;
}

//내지 구성 페이지 추가,수정 팝업
function popInfoReg(idx){
	var popInfoReg = window.open('/admin/organizer/option/pop_organizer_info_reg.asp?mode=modify&diaryid=' + idx,'popInfoReg','width=620,height=800,status=no,resizable=yes,scrollbars=yes')
	popInfoReg.focus();
}

//상세 내용 페이지 추가,수정 팝업
function popContReg(idx){
	alert('사용안함');
	var popContReg = window.open('/admin/organizer/pop_organizer_cont_reg.asp?mode=modify&organizerid=' + idx,'popContReg','width=620,height=800,resizable=yes,scrollbars=yes')
	popContReg.focus();
}


//알파 내용 페이지 추가,수정 팝업
function popalpha(idx){
	var popalpha = window.open('/admin/organizer/alpha_list.asp','popalpha','width=620,height=800,resizable=yes,scrollbars=yes')
	popalpha.focus();
}

//이벤트관리
function popeventReg(){
	var popeventReg = window.open('/admin/organizer/event.asp','popeventReg','width=1024,height=768,resizable=yes,scrollbars=yes')
	popeventReg.focus();
}
</script>


<!-- 상단 검색폼 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="refreshFrm" method="get">
<input type="hidden" name="menupos" value="<%= request("menupos") %>">
<input type="hidden" name="page" value="">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="80" bgcolor="<%= adminColor("gray") %>">검색조건</td>
	<td align="left">
		<% SelectList "cate",CateCode %>
		<select name="isusingbox">
		<option value=""<% if isusing = "" then response.write " selected"%>>사용여부</option>
		<option value="Y" <% if isusing = "Y" then response.write " selected"%>>Y</option>
		<option value="N" <% if isusing = "N" then response.write " selected"%>>N</option>	
		</select>
		브랜드:
		<% drawSelectBoxDesignerwithName "ebrand", sBrand %>
		<br>상품 코드:
		<input type="text" name="aitem" class="text" size="30" maxlength="50" value="<%= arrItemid %>"> 마지막 , 는 생략
	</td>
	<td width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onclick="refreshFrm.submit();">
	</td>
</tr>
</form>
</table>
<!-- 검색 끝 -->
<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding:10 0 10 0;">
<form name="frmarr" method="post" action="">
<input type="hidden" name="menupos" value="<%= request("menupos") %>">
<input type="hidden" name="mode" value="">
<tr>
	<td align="left">
		<input type="button" value="이미지관리" onclick="contents_option();" class="button">
		<input type="button" value="키워드관리" onclick="keyword_option();" class="button">
		<!--<input type="button" value="alpha배너관리" onclick="popalpha();" class="button">-->
		<!--<input type="button" value="이벤트관리" onclick="popeventReg();" class="button">-->
	</td>	
	<td align="right"><a href="javascript:popRegNew();"><img src="/images/icon_new_registration.gif" border="0"></a></td>
</tr>
</form>
</table>
<!-- 액션 끝 -->

<% If C_ADMIN_AUTH Then %>
	<table align="center" class="a">
	<tr>
		<td>
			매년 지난해의 판매 통계를 위해 백업테이블을 둠. 테이블 : [db_diary2010].[dbo].[diary_everyyear_for_statistic]<br>
			--insert into [db_diary2010].[dbo].[diary_everyyear_for_statistic]<br>
			select ItemID, '2012', 'o' from [db_diary2010].[dbo].[tbl_organizerMaster]<br>
			where isUsing = 'Y'<br>
			작업자는 매해 다이어리가 완전히 끝난 후 반드시 입력 필. 년도값은 2012~2013시즌일경우 2013.<br>
			다이어리인 경우는 'd', 오거나이저인 경우는 'o'.<br>
		</td>
	</tr>
	</table>
<% End If %>

<!-- 리스트 시작 -->
<table width="100%" border="0" align="center" class="a" cellpadding="0" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
	<% IF oip.FResultCount>0 Then %>
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15">
			검색결과 : <b><%= oip.FTotalCount %></b>
			&nbsp;
			페이지 : <b><%= page %>/ <%= oip.FTotalPage %></b>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td nowrap> 번호</td>
		<td nowrap> 구분 </td>
		<td nowrap> 이미지 </td>
		<td nowrap> 상품번호 </td>
		<td nowrap> 상품명 </td>      	
		<td nowrap> 사용여부 </td>
		<td nowrap> keyword </td>
		<td nowrap> 내지구성 </td>
		<!--<td>상품후기</td>-->
		<td nowrap> 관리 </td>
	</tr>

	<% For i =0 To  oip.FResultCount -1 %>
	<tr align="center" bgcolor="#FFFFFF">
		<td nowrap> <%= oip.FItemList(i).forganizerid %> </td>
		<td nowrap><% cateList "cate",oip.FItemList(i).FCateCode %> </td>
		<td nowrap>
			<img src="<%= db2html(oip.FItemList(i).ImgList) %>" width="40" height="40" border="0" style="cursor:pointer">
		</td>
		<td nowrap> <%= oip.FItemList(i).Fitemid %> </td>
		<td nowrap> <%= oip.FItemList(i).fitemname %> </td>      	
		<td><%= oip.FItemList(i).fisusing %> </td> 
		<td nowrap>
			<input type="button" class="button" value="등록" onClick="detail_view('<%= oip.FItemList(i).forganizerid %>');">
		</td>
		<td nowrap>
			<input type="button" class="button" value="등록" onclick="javascript:popInfoReg('<%= oip.FItemList(i).forganizerid %>');">	
			<!--<input type="button" class="button" value="등록" onclick="popInfoReg('<%= oip.FItemList(i).forganizerid %>');">-->
		</td>
		<!--<td align="center"><input type="button" class="button" value="등록" onclick="javascript:popRegeval(<%= oip.FItemList(i).Fitemid %>);"></td>-->
		<td nowrap>
			<input type="button" class="button" value="수정" onclick="popRegModi('<%= oip.FItemList(i).forganizerid %>');">
		</td>
	</tr>
	<% Next %>
<% else %>
		<tr bgcolor="#FFFFFF">
			<td colspan="3" align="center" class="page_link">[검색결과가 없습니다.]</td>
		</tr>
<% End IF %>
	<tr bgcolor="#FFFFFF">
		<td colspan="12" align="center" bgcolor="<%=adminColor("green")%>">
		
		<!-- 페이지 시작 -->
	    	<a href="?page=1&isusingbox=<%=isusing%>&cate=<%=catecode%>" onfocus="this.blur();"><img src="http://fiximage.10x10.co.kr/web2007/common/pprev_btn.gif" width="10" height="10" border="0"></a>
			<% if oip.HasPreScroll then %>
				<span class="list_link"><a href="?page=<%= oip.StartScrollPage-1 %>&isusingbox=<%=isusing%>&cate=<%=catecode%>">&nbsp;<img src="http://fiximage.10x10.co.kr/web2007/common/prev_btn.gif" width="10" height="10" border="0">&nbsp;</a></span>
			<% else %>
			&nbsp;<img src="http://fiximage.10x10.co.kr/web2007/common/prev_btn.gif" width="10" height="10" border="0">&nbsp;
			<% end if %>
			<% for i = 0 + oip.StartScrollPage to oip.StartScrollPage + oip.FScrollCount - 1 %>
				<% if (i > oip.FTotalpage) then Exit for %>
				<% if CStr(i) = CStr(oip.FCurrPage) then %>
				<span class="page_link"><font color="red"><b><%= i %>&nbsp;&nbsp;</b></font></span>
				<% else %>
				<a href="?page=<%= i %>&isusingbox=<%=isusing%>&cate=<%=catecode%>" class="list_link"><font color="#000000"><%= i %>&nbsp;&nbsp;</font></a>
				<% end if %>
			<% next %>
			<% if oip.HasNextScroll then %>
				<span class="list_link"><a href="?page=<%= i %>&isusingbox=<%=isusing%>&cate=<%=catecode%>">&nbsp;<img src="http://fiximage.10x10.co.kr/web2007/common/next_btn.gif" width="10" height="10" border="0">&nbsp;</a></span>
			<% else %>
			&nbsp;<img src="http://fiximage.10x10.co.kr/web2007/common/next_btn.gif" width="10" height="10" border="0">&nbsp;
			<% end if %>
			<a href="?page=<%= oip.FTotalpage %>&isusingbox=<%=isusing%>&cate=<%=catecode%>" onfocus="this.blur();"><img src="http://fiximage.10x10.co.kr/web2007/common/nnext_btn.gif" width="10" height="10" border="0"></a>
		<!-- 페이지 끝 -->
		
		</td>
	</tr>
</table>
<!-- 리스트 끝 -->

<% Set oip = nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
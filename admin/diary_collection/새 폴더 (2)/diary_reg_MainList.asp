<%@ language=vbscript %>
<% option explicit %>
<!-- include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/diary_collection/diary_collection_cls.asp" -->
<%

'// 추후 같은 포맷으로 진행시 이부분만 변경
const YearUse ="2008"


dim DiaryType,page
dim sql,i
dim pagesize

DiaryType=request("DiaryType")
page=request("page")
if page="" then page=1

pagesize =20

dim mdiary

set mdiary = new ClsDiary
mdiary.FYearUse =YearUse
mdiary.FDiaryType=DiaryType
mdiary.FCurrPage= page
mdiary.FPageSize=pagesize
mdiary.FScrollCount=10
mdiary.GetDiaryList




%>

<script type="text/javascript" language="javascript">
//분류별 검색
function FnSelDiaryType(varDiaryType){
	document.pagingFrm.page.value=1;
	document.pagingFrm.DiaryType.value=varDiaryType;
	document.pagingFrm.submit();
}
//페이지 이동
function FnPageMove(varPage){
	document.pagingFrm.page.value=varPage;
	document.pagingFrm.DiaryType.value='<%= DiaryType %>';
	document.pagingFrm.submit();
}
//신규 등록 팝업
function popRegNew(){
	window.open('/admin/diary_collection/diary_reg_new.asp?YearUse=<%= YearUse %>','newpop','width=450,height=400,status=yes')
}
//수정 팝업
function popRegModi(idx){
	window.open('/admin/diary_collection/diary_reg_edit.asp?mode=modify&idx=' + idx,'editpop','width=450,height=400')
}
//상세 내용 페이지 추가,수정 팝업
function popContReg(idx){
	window.open('/admin/diary_collection/diary_reg_Cont.asp?mode=modify&idx=' + idx,'contpop','width=620,height=800,resizable=yes,scrollbars=yes')
}
//내지 구성 페이지 추가,수정 팝업
function popInfoReg(idx){
	window.open('/admin/diary_collection/diary_reg_Info.asp?mode=modify&idx=' + idx,'infopop','width=620,height=800,status=no,resizable=yes,scrollbars=yes')
}
//관련 상품 추가,수정 팝업
function popLinkItemReg(idx){
	window.open('/admin/diary_collection/diary_reg_LinkItem.asp?idx=' + idx,'addpop','width=470,height=400,resizable=yes,scrollbars=yes')
}
//미리보기
function popPreview(idx){
	window.open('http://www.10x10.co.kr/diary_collection/diary_collection_prd.asp?diaryid=' + idx ,'prepop','width=750,height=600,resizable=yes,scrollbars=yes,menubar=yes,toolbar=yes,location=yes,status=yes')
}

//신상품 플래쉬 적용하기
function popNewItems(idx){
	window.open('/admin/diary_collection/diary_reg_NewItemList.asp?idx=' + idx ,'flashpop','width=750,height=600,resizable=yes,scrollbars=yes')
}
//베스트10 아이템 새로고침
function FnRefreshBest10(){
	window.open('http://test.10x10.co.kr/diary_collection_2007/do_diary_bestitem10.asp','bestpop','width=750,height=600,resizable=yes,scrollbars=yes');
}
// MD`s Pick 등록
function fnMdPickReg(){
	window.open('/admin/diary_collection/diary_mdspick.asp?yearUse=<%= yearUse %>','mdp','');
}
</script>
<!-- 상단 메뉴 -->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
	<tr height="10" valign="bottom">
        <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_02.gif"></td>
        <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<tr valign="top" style="padding : 0 0 10 0">
        <td background="/images/tbl_blue_round_04.gif"></td>
        <td align="right">
        	<table width="100%" border="0" cellpadding="0" cellspacing="0" class="a">
        		<tr>
        			<td>
        			<!--
        			<input type="button" class="button" onclick="popNewItems();" value="신상품 적용">
        			<input type="button" class="button" onclick="FnRefreshBest10();" value="베스트10새로고침">
        			-->
        			<input type="button" class="button" onclick="popRegNew();" value="신규 등록">
        			<select name="DiaryType"  onchange="FnSelDiaryType(this.value);">
						<option value="" 		 <% if DiaryType="" 		then response.write "selected"  %>>전체</option>
						<option value="illust" <% if DiaryType="illust" then response.write "selected"  %>>일러스트</option>
						<option value="photo"  <% if DiaryType="photo"  then response.write "selected"  %>>포토/명화</option>
						<option value="system" <% if DiaryType="system" then response.write "selected"  %>>시스템</option>
					</select>

        			</td>
        			<td align="right"><input type="button" class="button" value="md`s Pick 등록" onclick="fnMdPickReg();"></td>
        		</tr>
        	</table></td>
		<td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
	</tr>
</table>
<!-- 중간 메인부분 -->
<table width="100%" border="0" cellpadding="0" cellspacing="1"  class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr bgcolor="<%= adminColor("tabletop") %>">
		<td width="50" align="center">번호</td>
		<td width="60" align="center">구분</td>
		<td width="120" align="center">이미지</td>
		<td width="50" align="center">상품번호</td>
		<td align="center">상품명</td>
		<td width="70" align="center">상세 페이지</td>
		<td width="50" align="center">내지구성</td>
		<td width="50" align="center">관련상품</td>
	</tr>
	<!-- 리스트 표시 -->
	<% if mdiary.FResultCount<=0 then %>
		<!-- 없음 -->
	<% else %>
	<% for i = 0 to mdiary.FResultCount-1 %>

		<% if mdiary.FItemList(i).FIsusing="N" then %>
		<tr bgcolor="#ECECEC">
		<% else %>
		<tr bgcolor="#FFFFFF">
		<% end if %>

		<td align="center"><%= mdiary.FItemList(i).FIdx %></td>
		<td align="center"><%= mdiary.FItemList(i).StrDiaryTypeName %></td>
		<td align="center"><img src="<%= db2html(mdiary.FItemList(i).getListImgUrl) %>" width="100" height="100" border="1" onclick="popRegModi('<%= mdiary.FItemList(i).FIdx %>');" style="cursor:pointer"></td>
		<td align="center"><%= mdiary.FItemList(i).Fitemid %></td>
		<td align="center"><%= db2html(mdiary.FItemList(i).FItemName) %><span onclick="popPreview('<%= mdiary.FItemList(i).FIdx %>')" style="cursor:pointer"><font color="blue">[미리보기]</font></span></td>
		<td align="center"><span onclick="popContReg('<%= mdiary.FItemList(i).FIdx %>');" style="cursor:pointer">등록</span></td>
		<td align="center"><span onclick="popInfoReg('<%= mdiary.FItemList(i).FIdx %>');" style="cursor:pointer">등록</span></td>
		<td align="center"><span onclick="popLinkItemReg('<%= mdiary.FItemList(i).FIdx %>');" style="cursor:pointer">등록</span></td>
	</tr>
<% next %>
</table>
<% end if %>


<!-- 하단 페이징 시작 -->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
    <tr valign="bottom" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="center">
			<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
				<tr>
					<td align="center">
						<% if mdiary.HasPreScroll then %>
							<a href="javascript:FnPageMove('<%= mdiary.StartScrollPage-1 %>');">[pre]</a>
						<% else %>
							[pre]
						<% end if %>

						<% for i=0 + mdiary.StartScrollPage to mdiary.FScrollCount + mdiary.StartScrollPage - 1 %>
							<% if i>mdiary.FTotalpage then Exit for %>
							<% if CStr(page)=CStr(i) then %>
							<font color="red">[<%= i %>]</font>
							<% else %>
							<a href="javascript:FnPageMove('<%= i %>');">[<%= i %>]</a>
							<% end if %>
						<% next %>

						<% if mdiary.HasNextScroll then %>
							<a href="javascript:FnPageMove('<%= i %>');">[next]</a>
						<% else %>
							[next]
						<% end if %>
					</td>
				</tr>
			</table>
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    <tr valign="top" height="10">
        <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_08.gif"></td>
        <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
    </tr>
</table>
<!-- 페이징을 위한 폼 -->
<form name="pagingFrm" method="get" action="?">
<input type="hidden" name="page" value="" />
<input type="hidden" name="DiaryType" value="<%= DiaryType %>" />
<input type="hidden" name="menupos" value="<%= menupos %>" />
</form>
<!-- 페이징을 위한 폼 -->

<% set mdiary = nothing %>


<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->

<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/diary_collection/diary_collection_cls.asp" -->

<%
dim YearUse
YearUse = request("YearUse")

dim page,pagesize
page= request("page")
if page="" then page="1"
pagesize="5"

dim objMz,intLoop,arrList
set objMz = new ClsDiary

'objMz.FEvtType = evtType
objMz.FPageSize=pagesize
objMz.FCurrPage= page
'objMz.FEvtUsing = evtUsing
objMz.FScrollCount =10
arrList = objMz.getMagazineList()


%>
<script language="javascript">

function subchk(){

	if(document.regfrm.eventcode.value.length<1){
		document.regfrm.eventcode.focus();
		alert('이벤트 코드를 입력하셔야 합니다.');
		return false;
	}
	if(document.regfrm.multiname.value.length<1&&document.regfrm.leftname.value.length<1&&document.regfrm.powername.value.length<1){
		alert('이미지를 입력해 주세요');
		return false;
	}
	document.regfrm.submit();
}

function showimage(img){
	var pop = window.open('viewImage.asp?imageUrl='+img,'imgview','width=600,height=600,resizable=yes');
}

function FnPageMove(v){
	document.mainfrm.page.value=v;
	document.mainfrm.submit();
}

function fnDiaryMzReg(){
	document.location.href='/admin/diary_collection/pop_diary_magazine_reg.asp?yearUse=<%= yearUse %>';
}

function fnDiaryMzEdit(v){
	document.location.href='/admin/diary_collection/pop_diary_magazine_Edit.asp?yearUse=<%= yearUse %>&magazineid='+ v ;
}

document.domain = "10x10.co.kr";
window.resizeTo(500,500);
</script>
<!-- 상단 메뉴 -->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
	<tr height="10" valign="bottom">
        <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_02.gif"></td>
        <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<form name="mainfrm" method="get" action="">
	<input type="hidden" name="page" value="">
	<tr valign="top" style="padding : 0 0 10 0">
        <td background="/images/tbl_blue_round_04.gif"></td>
        <td align="right">
        	<input type="button" class="button" value="신규등록" onclick="fnDiaryMzReg();">
        	&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp
			<input type="button" class="button" value="검색" onclick="this.form.submit();">

		</td>
		<td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	</form>
</table>
<!-- 중단 내용 -->
<table width="100%" border="0" cellpadding="2" cellspacing="1" class="a" bgcolor="#9d9d9d">
	<tr bgcolor="#FFFFFF">
		<td align="center" width="40">번호</td>
		<td align="center" >제목</td>
		<td align="center" width="50">사용여부</td>
	</tr>
	<% if isArray(arrList) then %>
	<% for intLoop = 0 to Ubound(arrList,2) %>
	<tr bgcolor="#FFFFFF">
		<td align="center"><%= arrList(0,intLoop) %></td>
		<td align="center"><a href="javascript:fnDiaryMzEdit('<%= arrList(0,intLoop) %>');"><%= arrList(1,intLoop) %></a></td>
		<td align="center"><a href="javascript:fnDiaryMzEdit('<%= arrList(0,intLoop) %>');"><%= arrList(2,intLoop) %></a></td>
	</tr>
	<% next %>
	<% end if %>
</table>
<!-- 하단  시작 -->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
    <tr valign="bottom" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="center">
			<% if objMz.HasPreScroll then %>
				<a href="javascript:FnPageMove('<%= objMz.StartScrollPage-1 %>');">[pre]</a>
			<% else %>
				[pre]
			<% end if %>

			<% for intLoop=0 + objMz.StartScrollPage to objMz.FScrollCount + objMz.StartScrollPage - 1 %>
				<% if intLoop>objMz.FTotalpage then Exit for %>
				<% if CStr(page)=CStr(intLoop) then %>
				<font color="red">[<%= intLoop %>]</font>
				<% else %>
				<a href="javascript:FnPageMove('<%= intLoop %>');">[<%= intLoop %>]</a>
				<% end if %>
			<% next %>

			<% if objMz.HasNextScroll then %>
				<a href="javascript:FnPageMove('<%= intLoop %>');">[next]</a>
			<% else %>
				[next]
			<% end if %>
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    <tr valign="top" height="10">
        <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_08.gif"></td>
        <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
    </tr>
</table>
<% set objMz = nothing %>
</body>
</html>

<!-- #include virtual="/lib/db/dbclose.asp" -->

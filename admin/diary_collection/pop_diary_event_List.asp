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

dim evtUsing,BannerType
BannerType = request("BannerType")
if BannerType ="" then BannerType="multi"
evtUsing = request("evtUsing")
if evtUsing ="" then evtUsing ="Y"
dim page,pagesize
page= request("page")
if page="" then page="1"
pagesize="10"
dim objevt,intLoop,arrBan
set objevt = new ClsDiary
objevt.FBannerType = BannerType
objevt.FPageSize=pagesize
objevt.FCurrPage= page
objevt.FEvtUsing = evtUsing
objevt.FScrollCount =10
objevt.getBannerList




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

function fnDiaryEventReg(){
	document.location.href='/admin/diary_collection/pop_diary_event_reg.asp?yearUse=<%= yearUse %>';
}

function fnDiaryEventEdit(v){
	document.location.href='/admin/diary_collection/pop_diary_event_Edit.asp?yearUse=<%= yearUse %>&bannerid='+ v ;
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
			<input type="button" class="button" value="신규등록" onclick="fnDiaryEventReg();">
        	&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			적용 위치:
        	<select name="BannerType">

				<option value="multi" <% if BannerType="multi" then response.write "selected" %>>메인 멀티</option>
				<option value="left" <% if BannerType="left" then response.write "selected" %>>좌측 메뉴</option>
				<option value="power" <% if BannerType="power" then response.write "selected" %>>파워이벤트</option>
				<option value="today" <% if BannerType="today" then response.write "selected" %>>Today`s Diary</option>
				<option value="quiz" <% if BannerType="quiz" then response.write "selected" %>>Diary Quiz</option>
				<option value="dzone" <% if BannerType="dzone" then response.write "selected" %>>디자인존</option>
				<option value="tdayitem" <% if BannerType="tdayitem" then response.write "selected" %>>Today`s Item</option>
				<option value="evtmain" <% if BannerType="evtmain" then response.write "selected" %>>이벤트메인 배너</option>
				<option value="othermall_left" <% if BannerType="othermall_left" then response.write "selected" %>>[외부몰]좌측 메뉴</option>
				<option value="othermall_multi" <% if BannerType="othermall_multi" then response.write "selected" %>>[외부몰]메인 멀티</option>
				<option value="othermall_right" <% if BannerType="othermall_right" then response.write "selected" %>>[외부몰]우측 메뉴</option>
			</select>
			사용여부 <select name="evtUsing" size="1">
				<option value="Y" <% if evtUsing="Y" then response.write "selected" %>> Y &nbsp;&nbsp;</option>
				<option value="N" <% if evtUsing="N" then response.write "selected" %>> N &nbsp;&nbsp;</option>
			</select>
			<input type="button" class="button" value="검색" onclick="this.form.submit();">
		</td>
		<td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	</form>
</table>
<!-- 중단 내용 -->
<table width="100%" border="0" cellpadding="0" cellspacing="1" class="a" bgcolor="#9d9d9d">
	<tr bgcolor="#FFFFFF">
		<td align="center" width="50">번호</td>
		<td align="center" width="100">배너 TYPE</td>

		<td align="center" width="110">이미지</td>
		<!--<td align="center" width="50">사용여부</td>-->
	</tr>
	<% if objevt.FResultCount>0 then %>
	<% for intLoop = 0 to objevt.FResultCount-1 %>

	<% if objevt.FItemList(intLoop).FBannerUsing="Y" then %>
	<tr bgcolor="#FFFFFF">
	<% else %>
	<tr bgcolor="#ECECEC">
	<% end if %>



		<td align="center"><%= objevt.FItemList(intLoop).FBanneridx %></td>
		<td align="center"><a href="javascript:fnDiaryEventEdit('<%= objevt.FItemList(intLoop).FBanneridx %>')"><%= objevt.FItemList(intLoop).getBannerTypeStr %></a></td>

		<td align="center"><img src="<%= objevt.FItemList(intLoop).getBannerImgUrl %>" width="50" height="50" style="cursor:pointer" onclick="showimage('<%= objevt.FItemList(intLoop).getBannerImgUrl %>');"></td>
		<!--<td align="center"><%'= objevt.FItemList(intLoop).FBanneridx %></td> -->
	</tr>
	<% next %>
	<% end if %>
</table>
<!-- 하단  시작 -->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
    <tr valign="bottom" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="center">
			<% if objevt.HasPreScroll then %>
							<a href="javascript:FnPageMove('<%= objevt.StartScrollPage-1 %>');">[pre]</a>
						<% else %>
							[pre]
						<% end if %>

						<% for intLoop=0 + objevt.StartScrollPage to objevt.FScrollCount + objevt.StartScrollPage - 1 %>
							<% if intLoop>objevt.FTotalpage then Exit for %>
							<% if CStr(page)=CStr(intLoop) then %>
							<font color="red">[<%= intLoop %>]</font>
							<% else %>
							<a href="javascript:FnPageMove('<%= intLoop %>');">[<%= intLoop %>]</a>
							<% end if %>
						<% next %>

						<% if objevt.HasNextScroll then %>
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

<% set objevt = nothing %>
</body>
</html>

<!-- #include virtual="/lib/db/dbclose.asp" -->

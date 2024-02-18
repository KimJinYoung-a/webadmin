<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' History : 2007.11.09 한용민 개발
'			2008.06.18 한용민 수정
'###########################################################
%>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/sitemasterclass/othermall_main_event_rotationcls.asp"-->
<%

dim itemid , i ,page, malltype , isusing, research
	page = request("page")
	isusing = request("isusing")
	research = request("research")
	itemid = request("itemid")

	if (page = "") then
	        page = "1"
	end if
	if research="" and isusing="" then isusing="Y"

'==============================================================================
dim mdchoicerotate
set mdchoicerotate = new CMainMdChoiceRotate
	mdchoicerotate.FCurrPage = CInt(page)
	mdchoicerotate.FPageSize = 20
	mdchoicerotate.FRectIsUsing = isusing
	mdchoicerotate.FRectItemID = itemid
	mdchoicerotate.list

%>

<script language='javascript'>

function RefreshMainMdChoiceRotateEventRec(){
	if (confirm('메인 페이지에 적용 하시겠습니까?')){
		 var popwin = window.open('','refreshFrm','');
		 popwin.focus();
		 refreshFrm.target = "refreshFrm";
		 //refreshFrm.action = "http://uploadmain.10x10.co.kr/flash/link/MakeMainMdChoiceRotateFlash.asp" ;
		 refreshFrm.action = "<%=othermall%>/chtml/othermall_MakeMainMdChoiceRotateFlash.asp" ;
		 refreshFrm.submit();
	}
}

function NextPage(page){
	document.frm.page.value=page;
	document.frm.submit();
}

</script>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="refreshFrm" method="post">
	</form>
	<form name="frm" method="get" action="">
	<input type="hidden" name="page" value="1">
	<input type="hidden" name="research" value="on">
	<input type="hidden" name="menupos" value="<%= request("menupos") %>">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
		<td align="left">
			사용구분 :
			<select name="isusing" >
			<option value="" >전체
			<option value="Y" <% if isusing="Y" then response.write "selected" %> >사용함
			<option value="N" <% if isusing="N" then response.write "selected" %> >사용안함
			</select>
			상품번호 :
			<input type="text" name="itemid" value="<%= itemid %>" size=9 maxlength=9 onKeyDown = "javascript:onlyNumberInput()" style="IME-MODE: disabled" />
		</td>
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">

		</td>
	</tr>
	</form>
</table>
<!-- 검색 끝 -->

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<tr>
		<td align="left">
			사이트에 적용 <a href="javascript:RefreshMainMdChoiceRotateEventRec();">
			<img src="/images/refreshcpage.gif" width=19 align="absmiddle" border="0" alt="html만들기"></a>
		</td>
		<td align="right">
			<a href="othermall_main_md_recommend_flash_write.asp?mode=write&menupos=<%= menupos %>">
			<p align="right"><img src="/images/icon_new_registration.gif" width="75" border="0" align="absmiddle"></a>
		</td>
	</tr>
</table>
<!-- 액션 끝 -->

<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="vfrm" method="POST" action="">
<input type="hidden" name="page" value="<%=page%>">
<input type="hidden" name="menupos" value="<%= request("menupos") %>">
<input type="hidden" name="research" value="on">
<input type="hidden" name="sUsing" value="<%= isusing %>">
	<% if mdchoicerotate.FresultCount>0 then %>
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15">
			검색결과 : <b><%= mdchoicerotate.FTotalCount %></b>
			&nbsp;
			페이지 : <b><%= page %>/ <%= mdchoicerotate.FTotalPage %></b>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="30" align="center">ID</td>
		<td align="center" width="250">이미지</td>
		<td align="center">link정보</td>
		<td width="50" align="center">전시순서</td>
		<td width="100" align="center">등록일</td>
		<td width="50" align="center">사용유무</td>
		<td width="50" align="center">품절여부</td>
    </tr>
		<% for i=0 to mdchoicerotate.FresultCount-1 %>
	    <tr align="center" bgcolor="#FFFFFF">
		<td height="50" align="center">
			<input type="hidden" name="idx" value="<%= mdchoicerotate.FItemList(i).Fidx %>">
			<%= mdchoicerotate.FItemList(i).Fidx %>
		</td>
		<td align="center"><a href="othermall_main_md_recommend_flash_write.asp?mode=modify&idx=<%= mdchoicerotate.FItemList(i).Fidx %>&menupos=<%= menupos %>"><img src="<%= mdchoicerotate.FItemList(i).Fphotoimg %>" border=0 width="56"></a></td>
		<td height="50" align="left">
			<%= mdchoicerotate.FItemList(i).Flinkinfo %>
		</td>
		<td align="center">
			<input type="text" name="disporder" value="<%= mdchoicerotate.FItemList(i).FDisporder %>" size="3" style="text-align:right">
		</td>
		<td align="center">
			<%= FormatDateTime(mdchoicerotate.FItemList(i).Fregdate,2) %>
		</td>
		<td align="center">
			<select name="isusing">
				<option value="Y" <% if mdchoicerotate.FItemList(i).Fisusing="Y" then Response.Write "selected"%>>사용</option>
				<option value="N" <% if mdchoicerotate.FItemList(i).Fisusing="N" then Response.Write "selected"%>>불가</option>
			</select>
		</td>
		<td align="center">
			<% if mdchoicerotate.FItemList(i).IsSoldOut then %>
			<font color="red">품절</font>
			<% end if %>
		</td>
	    </tr>
		<% next %>

	<% else %>

		<tr bgcolor="#FFFFFF">
			<td colspan="3" align="center" class="page_link">[검색결과가 없습니다.]</td>
		</tr>

	<% end if %>

    <tr height="25" bgcolor="FFFFFF">
		<td colspan="15" align="center">
			<% if mdchoicerotate.HasPreScroll then %>
				<a href="javascript:NextPage('<%= mdchoicerotate.StarScrollPage-1 %>')">[pre]</a>
			<% else %>
				[pre]
			<% end if %>

			<% for i=0 + mdchoicerotate.StarScrollPage to mdchoicerotate.FScrollCount + mdchoicerotate.StarScrollPage - 1 %>
				<% if i>mdchoicerotate.FTotalpage then Exit for %>
				<% if CStr(page)=CStr(i) then %>
				<font color="red">[<%= i %>]</font>
				<% else %>
				<a href="javascript:NextPage('<%= i %>')">[<%= i %>]</a>
				<% end if %>
			<% next %>

			<% if mdchoicerotate.HasNextScroll then %>
				<a href="javascript:NextPage('<%= i %>')">[next]</a>
			<% else %>
				[next]
			<% end if %>
		</td>
	</tr>
	</form>
</table>

<%
set mdchoicerotate = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
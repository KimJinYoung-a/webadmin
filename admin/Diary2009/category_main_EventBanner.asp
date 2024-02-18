<%@ language=vbscript %>
<% option explicit %>
<%
'#######################################################
'	History	:  2009.04.18 한용민 2008프론트에서이동 2009용으로 변경
'#######################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/Diary2009/Classes/category_main_EventBannerCls.asp"-->

<%
'// 변수 선언
dim cdl, page, isusing, evtCd , cdm
	cdl = request("cdl")
	page = request("page")
	isusing = request("isusing")
	evtCd = request("evtCd")
	cdm		= request("cdm")

	if page="" then page=1
	if isusing="" then isusing="Y"

dim omd
	set omd = New CateEventBanner
	omd.FCurrPage = page
	omd.FPageSize=8
	omd.FRectCDL = cdl
	omd.FRectcdm = cdm
	omd.FRectEvtCD = evtCd
	omd.FRectIsusing = isusing
	omd.GetEventBannerList

dim i
%>

<script language='javascript'>

// 전체 체크/해제
function ckAll(){
	if(frm.idxArrTmp.length){
		for(i=0;i<frm.idxArrTmp.length;i++) {
			frm.idxArrTmp[i].checked=frm.ckall.checked;
		}
	}
	else {
		frm.idxArrTmp.checked=frm.ckall.checked;
	}
}

// 선택 체크
function CheckSelected(selc){
	if(frm.ckall.checked) {
		frm.ckall.checked=false;
		ckAll()
		selc.checked=true;
	}
}

// 선택 삭제여부 확인
function delitems(){
	var chk=0;
	if(frm.idxArrTmp.length) {
		for(i=0;i<frm.idxArrTmp.length;i++) {
			if(frm.idxArrTmp[i].checked)
				chk++;
		}
	}
	else {
		if(frm.idxArrTmp.checked)
			chk++;
	}

	if (chk==0){
		alert('선택아이템이 없습니다.');
		return;
	}
	
	
	if (confirm('선택 아이템을 삭제하시겠습니까?')){
		frm.mode.value="del";
		frm.action="doMainEventBanner.asp";
		frm.submit();
	}
}


// 전체 사용유무 변경
function changeUsing(upfrm){
	var chk=0;
	if(frm.idxArrTmp.length) {
		for(i=0;i<frm.idxArrTmp.length;i++) {
			if(frm.idxArrTmp[i].checked)
				chk++;
		}
	}
	else {
		if(frm.idxArrTmp.checked)
			chk++;
	}

	if (chk==0){
		alert('선택아이템이 없습니다.');
		return;
	}
	
	if (upfrm.allusing.value=='Y'){
		var ret = confirm('선택 아이템을 사용함으로 변경합니다');
	} else {
		var ret = confirm('선택 아이템을 사용안함 으로  변경합니다');
	}
	
	if (ret){
		
		upfrm.mode.value="changeUsing";
		upfrm.action="doMainEventBanner.asp";
		upfrm.submit();

	}
}

// 감성채널 파일삭제
function delCategoryEventBanner(){
    if(!frm.cdl.value) {
    	alert("적용할 카테고리를 선택해주십시오.");
    }
   	else {
	    if (confirm('적용하시겠습니까?')){
			 var popwin = window.open('','refreshFrm','');
			 popwin.focus();
			 refreshFrm.target = "refreshFrm";
			 refreshFrm.action = "<%=wwwURL%>/chtml/diary/make_category_main_EventBanner_del.asp?cdl=" + frm.cdl.value+"&cdm="+frm.cdm.value;
			 refreshFrm.submit();
	    }
	}
}

// 이벤트 배너 페이지 적용여부 확인
function RefreshCategoryEventBanner(){
    if(!frm.cdl.value) {
    	alert("적용할 카테고리를 선택해주십시오.");
    }
	else if (frm.cdl.value == '110'){
		if (frm.cdm.value==''){
			alert('감성채널은 중카테고리를 선택해야만 합니다');			
			return;
		}else{
		    if (confirm('감성채널 적용하시겠습니까?')){
				 var popwin = window.open('','refreshFrm','');
				 popwin.focus();
				 refreshFrm.target = "refreshFrm";
				 refreshFrm.action = "<%=wwwURL%>/chtml/diary/make_category_main_EventBanner.asp?cdl=" + frm.cdl.value + "&cdm=" + frm.cdm.value;
				 refreshFrm.submit();
		    }		
		}
	}
   	else {
	    if (confirm('적용하시겠습니까?')){
			 var popwin = window.open('','refreshFrm','');
			 popwin.focus();
			 refreshFrm.target = "refreshFrm";
			 refreshFrm.action = "<%=wwwURL%>/chtml/diary/make_category_main_EventBanner.asp?cdl=" + frm.cdl.value + "&cdm=" + frm.cdm.value;
			 refreshFrm.submit();
	    }
	}
}

// 내용 수정
function viewPage(idx)
{
	frm.mode.value="edit";
	frm.page.value=<%=page%>;
	frm.idx.value=idx;
	frm.action="category_main_EventBanner_input.asp";
	frm.submit();
}

function changecontent()
{
	document.frm.submit();

}

</script>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="refreshFrm" method="post"></form>
<form name="frm" method="get" action="">
<input type="hidden" name="page" value="1">
<input type="hidden" name="mode" value="">
<input type="hidden" name="evtid" value="">
<input type="hidden" name="idx" value="">
<input type="hidden" name="idxarr" value="">
<input type="hidden" name="menupos" value="<%= request("menupos") %>">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
		<td align="left">
			카테고리선택 : 
<select class='select' name="cdl">
<option value='010' selected>디자인문구</option>
</select>
<select class='select' name="cdm">
<option value='' <% If cdm = "" Then Response.Write "selected" End If %>>전체다이어리</option>
<option value='10' <% If cdm = "10" Then Response.Write "selected" End If %>>심플다이어리</option>
<option value='20' <% If cdm = "20" Then Response.Write "selected" End If %>>일러스트다이어리</option>
<option value='30' <% If cdm = "30" Then Response.Write "selected" End If %>>캐릭터다이어리</option>
<option value='40' <% If cdm = "40" Then Response.Write "selected" End If %>>포토다이어리</option>
</select>
			<% if cdl="110" then DrawSelectBoxCategoryMid "cdm", cdl, cdm %>
			<a href="javascript:RefreshCategoryEventBanner()"><img src="/images/refreshcpage.gif" width="19" height="23" border="0" align="absmiddle">실서버적용</a> 
			<% if cdl="110" then %>
				<input type="button" value="실서버파일삭제" onclick="delCategoryEventBanner()" class="button">
			<% end if %>
			<br>사용유무 : <select name="isusing"><option value="Y">Yes</option><option value="N">No</option></select>
			이벤트코드 : <input type="text" name="evtCd" value="<%=evtCd%>" size="6">
			<script language="javascript">
				document.frm.isusing.value="<%=isusing%>";
			</script>			
		</td>	
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="검색" onClick="frm.submit();">
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">	
		</td>
	</tr>
</table>
<!-- 검색 끝 -->

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<tr>
		<td align="left">
			<select name="allusing"><option value="Y">선택 -> Y</option><option value="N">선택 ->N </option></select><input type="button" class="button" value="적용" onclick="changeUsing(frm);">
			<% if cdl<>"" then %>
				<input type="button" value="선택아이템삭제" onclick="delitems();" class="button">
			<% end if %>		
		</td>
		<td align="right">		
			<input type="button" value="아이템 추가" onclick="self.location='/admin/diary2009/category_main_EventBanner_input.asp?mode=add&cdl=<%= cdl %>&menupos=<%= menupos %>'" class="button">					
		</td>
	</tr>
</table>
<!-- 액션 끝 -->

<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<% if omd.FresultCount>0 then %>
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15">
			검색결과 : <b><%= omd.FTotalCount %></b>
			&nbsp;
			페이지 : <b><%= page %>/ <%= omd.FTotalPage %></b>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="50" align="center"><input type="checkbox" name="ckall" onclick="ckAll()"></td>
		<td width="100" align="center">카테고리명</td>
		<td width="150" align="center">이벤트명</td>
		<td align="center">이미지</td>
		<td width="50" align="center">표시순서</td>
		<td width="50" align="center">사용유무</td>
		<td width="80" align="center">등록일</td>
    </tr>
	<% for i=0 to omd.FResultCount-1 %>
		
    <% if omd.FItemList(i).fisusing = "Y" then %>
	    <tr align="center" bgcolor="#FFFFFF">
	    <% else %>    
	    <tr align="center" bgcolor="#FFFFaa">
		<% end if %>
		<td align="center"><input type="checkbox" name="idxArrTmp" value="<%= omd.FItemList(i).fidx %>" onclick="CheckSelected(this)"></td>
		<td align="center">
			<%= omd.FItemList(i).Fcode_nm %>
			<% if omd.FItemList(i).fcdm <> "" then %>
				(<%=omd.FItemList(i).Fcode_nm_mid%>)
			<% end if %>
		</td>
		<td align="center"><a href="javascript:viewPage(<%= omd.FItemList(i).fidx %>);"><%= "[" & omd.FItemList(i).Fevt_code & "] " & omd.FItemList(i).Fevt_name %></a></td>
		<td align="center"><img src="<%= omd.FItemList(i).Fevt_bannerimg %>" width="200" border="0"></td>
		<td align="center"><%= omd.FItemList(i).Fviewidx %></td>
		<td align="center"><%= omd.FItemList(i).Fisusing %></td>
		<td align="center"><%= FormatDateTime(omd.FItemList(i).Fregdate,2) %></td>
    </tr>   

	<% next %>
	<% else %>
		<tr bgcolor="#FFFFFF">
			<td colspan="3" align="center" class="page_link">[검색결과가 없습니다.]</td>
		</tr>
	<% end if %>
    <tr height="25" bgcolor="FFFFFF">
		<td colspan="15" align="center">
			<% if omd.HasPreScroll then %>
				<a href="?page=<%= omd.StarScrollPage-1 %>&menupos=<%= menupos %>&isusing=<%=isusing%>&cdl=<%=cdl%>&evtCd=<%=evtCd%>">[pre]</a>
			<% else %>
				[pre]
			<% end if %>
		
			<% for i=0 + omd.StarScrollPage to omd.FScrollCount + omd.StarScrollPage - 1 %>
				<% if i>omd.FTotalpage then Exit for %>
				<% if CStr(page)=CStr(i) then %>
				<font color="red">[<%= i %>]</font>
				<% else %>
				<a href="?page=<%= i %>&menupos=<%= menupos %>&isusing=<%=isusing%>&cdl=<%=cdl%>&evtCd=<%=evtCd%>">[<%= i %>]</a>
				<% end if %>
			<% next %>
		
			<% if omd.HasNextScroll then %>
				<a href="?page=<%= i %>&menupos=<%= menupos %>&isusing=<%=isusing%>&cdl=<%=cdl%>&evtCd=<%=evtCd%>">[next]</a>
			<% else %>
				[next]
			<% end if %>
		</td>
	</tr>
</form>	
</table>

<%
set omd = Nothing
%>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
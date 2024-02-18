<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  기프트
' History : 2014.03.19 한용민 생성
' History : 2014.10.31 유태욱 mtitle 추가
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/sitemaster/gift/giftday_cls.asp"-->
<%
dim research, title, mtitle, isusing, page, masteridx, cgiftday, i
	title	= requestcheckvar(request("title"),128)
	mtitle	= requestcheckvar(request("mtitle"),128)
	page	= requestcheckvar(request("page"),10)
	isusing	= requestcheckvar(request("isusing"),1)
	research	= requestcheckvar(request("research"),2)
	menupos	= requestcheckvar(request("menupos"),10)
	masteridx	= requestcheckvar(request("masteridx"),10)
	
If page = ""	Then page = 1
if research ="" and isusing="" then isusing = "Y"

SET cgiftday = new Cgiftday_list
	cgiftday.FCurrPage		= page
	cgiftday.FPageSize		= 50
	cgiftday.Frecttitle		= title
	cgiftday.Frectmtitle		= mtitle
	cgiftday.Frectisusing		= isusing
	cgiftday.Frectmasteridx		= masteridx
	cgiftday.getgiftday_master
%>

<script type='text/javascript'>

var ichk;
ichk = 1;

function jsChkAll(){
	var frm, blnChk;
	frm = document.fitem;
	if(!frm.chkI) return;
	if ( ichk == 1 ){
		blnChk = true;
		ichk = 0;
	}else{
		blnChk = false;
		ichk = 1;
	}
	for (var i=0;i<frm.elements.length;i++){
		var e = frm.elements[i];
		if ((e.type=="checkbox")) {
			e.checked = blnChk ;
		}
	}
}

// 이미지 클릭시 원본 크기로 팝업 보기
function doImgPop(img){
	img1= new Image();
	img1.src=(img);
	imgControll(img);
}

function imgControll(img){
	if((img1.width!=0)&&(img1.height!=0)){
		viewImage(img);
	}else{
		controller="imgControll('"+img+"')";
		intervalID=setTimeout(controller,20);
	}
}

function viewImage(img){
	W=img1.width;
	H=img1.height;
	O="width="+W+",height="+H+",scrollbars=yes";
	imgWin=window.open("","",O);
	imgWin.document.write("<html><head><title>:*:*:*: 이미지상세보기 :*:*:*:*:*:*:</title></head>");
	imgWin.document.write("<body topmargin=0 leftmargin=0>");
	imgWin.document.write("<img src="+img+" onclick='self.close()' style='cursor:pointer;' title ='클릭하시면 창이 닫힙니다.'>");
	imgWin.document.close();
}

function giftdayedit(masteridx){
	var giftdayedit = window.open('/admin/sitemaster/gift/day/giftday_edit.asp?masteridx='+masteridx+'&menupos=<%=menupos%>','giftdayedit','width=1024,height=768,scrollbars=yes,resizable=yes');
	giftdayedit.focus();
}

function giftdaywinner(masteridx){
	var giftdaywinner = window.open('/admin/sitemaster/gift/day/giftdaywinner.asp?masteridx='+masteridx+'&menupos=<%=menupos%>','giftdaywinner','width=1024,height=768,scrollbars=yes,resizable=yes');
	giftdaywinner.focus();
}

function gosubmit(page){
    var frm = document.frm;
    frm.page.value=page;
	frm.submit();
}

function jsSetItem(idx){
	var popitem;
	popitem = window.open('/admin/sitemaster/gift/day/giftday_item.asp?idx='+idx,'popitem','width=920,height=600,scrollbars=yes,resizable=yes');
	popitem.focus();
}

</script>

<!-- 검색 시작 -->
<form name="frm" method="get" action="" style="margin:0px;">
<input type="hidden" name="page" value="<%=page%>">
<input type="hidden" name="research" value="on">
<input type="hidden" name="menupos" value="<%= menupos %>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		* 번호 : <input type="text" name="masteridx" value="<%=masteridx%>" size="10" maxlength="10" class="text">
		&nbsp;&nbsp;
		* 제목 : <input type="text" name="title" value="<%=title%>" size="40" maxlength="40" class="text">
		&nbsp;&nbsp;
		* 사용유무 :
		<% drawSelectBoxUsingYN "isusing", isusing %>	
	</td>	
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="gosubmit('');">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
	</td>
</tr>
</table>
</form>
<!-- 검색 끝 -->

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
	</td>
	<td align="right">
		<input type="button" value="주제신규등록" class="button" onclick="giftdayedit('');">
	</td>
</tr>
</table>
<!-- 액션 끝 -->

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="fitem" method="post" style="margin:0px;">
<input type="hidden" name="sortnoarr" value="">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="20">
		검색결과 : <b><%=cgiftday.FTotalCount %></b>
		&nbsp;
		페이지 : <b><%= page %>/ <%= cgiftday.FTotalPage %></b>		
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="30">
	<!--<td><input type="checkbox" name="chkA" onClick="jsChkAll();"></td>-->
	<td>번호</td>
	<td>WWW<Br>리스트탑</td>
	<td>제목</td>
	<td>모바일제목</td>
	<td>기간</td>
	<td>사용<Br>여부</td>
	<td>사연수</td>
	<td>비고</td>
</tr>
<% if cgiftday.fresultcount > 0 then %>
<% For i = 0 to cgiftday.fresultcount -1 %>
<% if cgiftday.FItemList(i).fisusing="Y" then %>
	<tr height="25" bgcolor="FFFFFF"  align="center">
<% else %>
	<tr height="25" bgcolor="f1f1f1"  align="center">
<% end if %>	
	<!--<td align="center"><input type="checkbox" name="chkI" onClick="AnCheckClick(this);"  value="<%= cgiftday.FItemList(i).fmasteridx %>"></td>-->
	<td align="center"><%= cgiftday.FItemList(i).fmasteridx %></td>
	<td align="center">
		<img src="<%=cgiftday.FItemList(i).flisttopimg_w%>" width="50" height="50" title="클릭하시면 원본크기로 보실 수 있습니다." style="cursor: pointer;" onclick="doImgPop('<%=cgiftday.FItemList(i).flisttopimg_w%>')"/>
	</td>
	<td align="center"><%= ReplaceBracket(cgiftday.FItemList(i).ftitle) %></td>
	<td align="center"><%= ReplaceBracket(cgiftday.FItemList(i).fmtitle) %></td>
	<td align="center"><%= left(cgiftday.FItemList(i).fstartdate,10) %> - <%= left(cgiftday.FItemList(i).fenddate,10) %></td>
	<td><%=cgiftday.FItemList(i).FIsusing%></td>
	<td><%=cgiftday.FItemList(i).fdetailcount%></td>
	<td>
		<input type="button" onClick="giftdayedit('<%=cgiftday.FItemList(i).fmasteridx%>');" value="수정" class="button">
		<input type="button" onClick="giftdaywinner('<%=cgiftday.FItemList(i).fmasteridx%>');" value="참여리스트" class="button">
		<input type="button" class="button" value="상품확인[<%= cgiftday.FItemList(i).Fitemcnt %>]" onclick="jsSetItem('<%= cgiftday.FItemList(i).fmasteridx %>','0');"/>
	</td>	
</tr>
<% Next %>
<tr height="25" bgcolor="FFFFFF" >
	<td colspan="15" align="center">
       	<% If cgiftday.HasPreScroll Then %>
			<span class="cgiftday_link"><a href="javascript:gosubmit('<%= cgiftday.StartScrollPage-1 %>');">[pre]</a></span>
		<% Else %>
		[pre]
		<% End If %>
		<% For i = 0 + cgiftday.StartScrollPage to cgiftday.StartScrollPage + cgiftday.FScrollCount - 1 %>
			<% If (i > cgiftday.FTotalpage) Then Exit for %>
			<% If CStr(i) = CStr(cgiftday.FCurrPage) Then %>
			<span class="page_link"><font color="red"><b><%= i %></b></font></span>
			<% Else %>
			<a href="javascript:gosubmit('<%= i %>');" class="cgiftday_link"><font color="#000000"><%= i %></font></a>
			<% End if %>
		<% Next %>
		<% If cgiftday.HasNextScroll Then %>
			<span class="cgiftday_link"><a href="javascript:gosubmit('<%= i %>');">[next]</a></span>
		<% Else %>
		[next]
		<% End If %>
	</td>
</tr>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="20" align="center" class="page_link">[검색결과가 없습니다.]</td>
	</tr>
<% end if %>
</form>
</table>

<% 
SET cgiftday = nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
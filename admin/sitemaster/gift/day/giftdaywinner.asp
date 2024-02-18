<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  기프트
' History : 2014.03.19 한용민 생성
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
dim research, isusing, page, masteridx, detailidx, cgiftday, i, userid, order
	userid	= requestcheckvar(request("userid"),32)
	page	= requestcheckvar(request("page"),10)
	isusing	= requestcheckvar(request("isusing"),1)
	research	= requestcheckvar(request("research"),2)
	menupos	= requestcheckvar(request("menupos"),10)
	masteridx	= requestcheckvar(request("masteridx"),10)
	detailidx	= requestcheckvar(request("detailidx"),10)
	order	= requestcheckvar(request("order"),32)

If page = ""	Then page = 1
if research ="" and isusing="" then isusing = "Y"
if order="" then order="new"

if masteridx="" then
	Response.Write "<script language='javascript'>alert('주제번호가 없습니다.');</script>"
	dbget.close()	:	response.End
end if

SET cgiftday = new Cgiftday_list
	cgiftday.FCurrPage		= page
	cgiftday.FPageSize		= 50
	cgiftday.Frectuserid		= userid
	cgiftday.Frectisusing		= isusing
	cgiftday.Frectmasteridx		= masteridx
	cgiftday.Frectdetailidx		= detailidx
	cgiftday.frectorder = order
	
	if masteridx<>"" then
		cgiftday.getgiftday_winner
	end if
%>

<script language="javascript">

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


function gosubmit(page){
    var frm = document.frm;
    frm.page.value=page;
	frm.submit();
}

function delwinnder(detailidx){
	if (detailidx==""){
		alert('구분자가 없습니다.');
		return;
	}
	
	if (confirm('사연을 삭제 하시겠습니까?')){
		frmreg.mode.value="del";
		frmreg.detailidx.value=detailidx;
		frmreg.action="/admin/sitemaster/gift/day/giftdaywinner_proc.asp";
		frmreg.submit();
	}
}

</script>

<!-- 검색 시작 -->
<form name="frm" method="post" action="" style="margin:0px;">
<input type="hidden" name="page" value="<%=page%>">
<input type="hidden" name="research" value="on">
<input type="hidden" name="masteridx" value="<%=masteridx%>">
<input type="hidden" name="menupos" value="<%= menupos %>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		* 주제번호 : <%=masteridx%>
		&nbsp;&nbsp;
		* 사연번호 : <input type="text" name="detailidx" value="<%=detailidx%>" size="10" maxlength="10" class="text">		
		&nbsp;&nbsp;
		* 고객ID : <input type="text" name="userid" value="<%=userid%>" size="20" maxlength="32" class="text">
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
<!-- 검색 끝 -->

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
	</td>
	<td align="right">
	</td>
</tr>
</table>
<!-- 액션 끝 -->

<Br>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="20">
		검색결과 : <b><%=cgiftday.FTotalCount %></b>
		&nbsp;
		페이지 : <b><%= page %>/ <%= cgiftday.FTotalPage %></b>
		<p align="right">
			* 정렬 : 
			<select name="order" onchange="gosubmit('');">
				<option value="new" <% if order="new" then response.write " selected" %>>최신순</option>
				<option value="comment" <% if order="comment" then response.write " selected" %>>코맨수수</option>
				<option value="join" <% if order="join" then response.write " selected" %>>참여횟수</option>
			</select>
		</p>
	</td>
</tr>
</form>

<form name="frmreg" method="post" action="" style="margin:0px;">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="detailidx">
<input type="hidden" name="mode">
</form>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="30">
	<!--<td><input type="checkbox" name="chkA" onClick="jsChkAll();"></td>-->
	<td>사연<br>번호</td>
	<td>회원ID</td>
	<td>상품<br>이미지</td>
	<td>나이</td>
	<td>회원등급</td>
	<td>작성일</td>
	<td>내용</td>
	<td>코맨트<br>수</td>
	<td>참여<br>횟수</td>
	<td>비고</td>
</tr>
<% if cgiftday.fresultcount > 0 then %>
<% For i = 0 to cgiftday.fresultcount -1 %>
<% if cgiftday.FItemList(i).fisusing="Y" then %>
	<tr height="25" bgcolor="FFFFFF"  align="center">
<% else %>
	<tr height="25" bgcolor="f1f1f1"  align="center">
<% end if %>	
	<!--<td align="center"><input type="checkbox" name="chkI" onClick="AnCheckClick(this);"  value="<%= cgiftday.FItemList(i).fdetailidx %>"></td>-->
	<td align="center"><%= cgiftday.FItemList(i).fdetailidx %></td>
	<td align="center"><%= cgiftday.FItemList(i).fuserid %></td>
	<td align="center">
		<img src="<%=cgiftday.FItemList(i).fimagesmall%>" width="50" height="50" title="클릭하시면 원본크기로 보실 수 있습니다." style="cursor: pointer;" onclick="doImgPop('<%=cgiftday.FItemList(i).fimagesmall%>')"/>
	</td>
	<td align="center"><%= cgiftday.FItemList(i).fage %></td>
	<td align="center"><%= getUserLevelStrByDate(cgiftday.FItemList(i).fuserlevel, left(cgiftday.FItemList(i).fregdate,10)) %></td>
	<td align="center"><%= left(cgiftday.FItemList(i).fregdate,10) %></td>
	<td width=300><%=cgiftday.FItemList(i).fcontents%></td>
	<td><%=cgiftday.FItemList(i).fcommentcount%></td>
	<td><%=cgiftday.FItemList(i).fjoincount%></td>
	<td>
		<input type="button" onclick="delwinnder('<%= cgiftday.FItemList(i).fdetailidx %>')" value="삭제" class="button">
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
</table>

<% 
SET cgiftday = nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
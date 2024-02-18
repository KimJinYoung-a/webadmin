<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 업체 입점문의
' History : 서동석 생성
'			2008.09.01 한용민 수정
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/classes/board/companyrequestcls.asp" -->
<!-- #include virtual="/lib/classes/displaycate/displaycateCls.asp"-->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<%
dim i, j, arrUrl, commmode, page,gubun, onlymifinish, research, searchkey,catevalue,dispCate,maxDepth, ipjumYN, comreqID
dim companyrequest
	commmode=requestCheckvar(request("commmode"),10)
	page	 	= requestCheckvar(request("pg"),10)
	gubun 		= requestCheckvar(request("gubun"),2)
	onlymifinish= requestCheckvar(request("onlymifinish"),3)
	research 	= requestCheckvar(request("research"),3)
	searchkey 	= requestCheckvar(request("searchkey"),32)
	catevalue	= requestCheckvar(request("catevalue"),3)
	ipjumYN		= requestCheckvar(request("ipjumYN"),1)
	comreqID 	= requestCheckvar(request("id"),10)
	dispCate		= requestCheckVar(Request("disp"),16) 
	maxDepth		= 2

if research="" and onlymifinish="" then onlymifinish="on"

'// 기본값으로 입점의뢰서
if gubun="" then gubun="01"
if (page = "") then page = "1"


set companyrequest = New CCompanyRequest
	companyrequest.read(comreqID)

%> 
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript">

function delreq(){
	if (confirm("삭제하시겠습니까?") ==true)
		frm.mode.value="reqdel"; 
		frm.action="/admin/board/upche/req_act.asp";
		frm.submit();
}

function SubmitForm(){
	if (confirm("처리상태를 완료로 전환합니까?") == true) { document.f.submit(); }
}
function catesubmit(){

	if (confirm("카테고리를 변경 합니다.") ==true)
		frm.mode.value="chcate"; 
		frm.disp.value=f.disp.value; 
		frm.action="/admin/board/upche/req_act.asp";
		frm.submit();
}
function sellsubmit(){

	if (confirm("판매형식을 변경합니다.") ==true)
		frm.mode.value="chsell";
		frm.sellgubun.value=f.sellgubun.value;
		frm.action="/admin/board/upche/req_act.asp";
		frm.submit();
}
function ipjumYNsubmit(){

	if(confirm("입점여부 선택합니다.") ==true)
		frm.mode.value="ipjum";
		frm.ipjumYN.value=f.ipjumYN.value;
		frm.action="/admin/board/upche/req_act.asp";
		frm.submit();
}

function sendmail(){
    var ireqmail = "<%= replace(replace(replace(trim(companyrequest.results(0).email),"<br>",""),vbcrlf,""),"\n","") %>";

    if (ireqmail.length<2){
        alert('메일주소가 올바르지 않습니다.');
        return;
    }
    
	if(confirm("메일을 보내시겠습니까?.") ==true)
	frmmail.submit();
}

function MovePage(page){
	frm.pg.value=page;
	frm.research.value="<%=research %>";
	frm.gubun.value="<%=gubun%>";
	frm.onlymifinish.value="<%=onlymifinish%>";
	frm.catevalue.value="<%=catevalue%>";
	frm.ipjumYNvalue="<%=ipjumYN%>";
	frm.searchkey.value="<%=searchkey%>";
	frm.action="/admin/board/upche/req_list.asp";
	frm.submit();
}
function editcomm(){
	frm.commmode.value="edit";
	frm.id.value="<%= companyrequest.results(0).id %>";
	frm.user.value="<%= session("ssBctCname") %>";
	frm.action="/admin/board/upche/req_view.asp";
	frm.submit();
}
function savecomm(){
	frm.mode.value="comm";
	frm.id.value="<%= companyrequest.results(0).id %>";
	frm.user.value="<%= session("ssBctCname") %>";
	frm.comment.value=commfrm.comment.value;
	frm.action="/admin/board/upche/req_act.asp";
	frm.submit();
	}

function AddNewBrand(){
	var cate1 = document.f.disp.value;
	var popwin = window.open("/admin/member/addnewbrand_step1.asp?pcuserdiv=9999_02&companyno=<%= db2html(companyrequest.results(0).license_no) %>&hp=<%= db2html(companyrequest.results(0).hp) %>&email=<%= db2html(replace(replace(replace(trim(companyrequest.results(0).email),"<br>",""),vbcrlf,""),"\n","")) %>&cd1=<%= left(companyrequest.results(0).dispcate,3) %>&cate1="+cate1,"addnewbrand2","width=800 height=580 scrollbars=yes resizable=yes");
	popwin.focus();
}

</script>

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td align="left">
		※고객센타 - 업체상담게시판
	</td>
</tr>
</table>
<!-- 액션 끝 -->

<!-- 업체정보 시작 -->
<form method="post" name="f" action="/admin/board/upche/req_act.asp" onsubmit="return false" style="margin:0px;">
<input type="hidden" name="mode" value="finish">
<input type="hidden" name="menupos" value="menupos">
<input type="hidden" name="id" value="<%= companyrequest.results(0).id %>">
<table width="100%" align="center" cellpadding="5" cellspacing="1" bgcolor="black" class="a">
<tr bgcolor="FFFFFF">
	<td colspan=5><b><font color="blue">협력업체 관리 선정기준</font></b></td>	
</tr>
<tr bgcolor="FFFFFF">
	<td width="100" bgcolor="<%= adminColor("gray") %>" align="center">회사명</td>	
	<td><%= db2html(companyrequest.results(0).companyname) %></td>	
	<td width="100" bgcolor="<%= adminColor("gray") %>" align="center">대표자명</td>			
	<td><%= db2html(companyrequest.results(0).chargename) %></td>				
</tr>

<tr bgcolor="FFFFFF">
	<td bgcolor="<%= adminColor("gray") %>" align="center">주소</td>	
	<td><%= db2html(companyrequest.results(0).address) %></td>
	<td bgcolor="<%= adminColor("gray") %>" align="center">구매고객</td>	
	<td>
		<%= db2html(companyrequest.results(0).cur_target) %>
	</td>							
</tr>
	
<tr bgcolor="FFFFFF">
	<td bgcolor="<%= adminColor("gray") %>" align="center">성명</td>			
	<td><%= db2html(companyrequest.results(0).chargename) %></td>	
	<td bgcolor="<%= adminColor("gray") %>" align="center">직책(부서명)</td>			
	<td><%= db2html(companyrequest.results(0).chargeposition) %></td>			
</tr>
<tr bgcolor="FFFFFF">
	<td bgcolor="<%= adminColor("gray") %>" align="center">Tel</td>			
	<td><%= db2html(companyrequest.results(0).phone) %></td>	
	<td bgcolor="<%= adminColor("gray") %>" align="center">H.P</td>			
	<td><%= db2html(companyrequest.results(0).hp) %></td>	
</tr>
<tr bgcolor="FFFFFF">
	<td bgcolor="<%= adminColor("gray") %>" align="center">사업자등록번호</td>	
	<td><%= db2html(companyrequest.results(0).license_no) %></td>
	<td bgcolor="<%= adminColor("gray") %>" align="center">이메일</td>			
	<td><a href="mailto:<%= db2html(replace(replace(replace(trim(companyrequest.results(0).email),"<br>",""),vbcrlf,""),"\n","")) %>"><%= db2html(replace(replace(replace(trim(companyrequest.results(0).email),"<br>",""),vbcrlf,""),"\n","")) %></a></td>	
</tr>
<tr bgcolor="FFFFFF">
	<td bgcolor="<%= adminColor("gray") %>" align="center">업태</td>	
	<td>
		<% 
		if companyrequest.results(0).Service <> "" then
			if left(companyrequest.results(0).Service,1) <> 0 then response.write "제조. "
			if mid(companyrequest.results(0).Service,3,1) <> 0 then response.write "도매. "
			if mid(companyrequest.results(0).Service,5,1) <> 0 then response.write "소매. "	 
			if mid(companyrequest.results(0).Service,7,1) <> 0 then response.write "수출. "
			if mid(companyrequest.results(0).Service,9,1) <> 0 then response.write "서비스. "
			if mid(companyrequest.results(0).Service,11,1) <> 0 then response.write "수입. "	
			if right(companyrequest.results(0).Service,1) <> 0 then response.write "기타. "
		end if
		%>
	</td>
		
	<td bgcolor="<%= adminColor("gray") %>" align="center">주품목</td>	
	<td>
		<% Drawcatelarge "catelargebox",companyrequest.results(i).cd1 %>(<% Drawcatemid companyrequest.results(0).cd1,"catemidbox",companyrequest.results(0).cd2 %>)
	</td>				
</tr>	
<tr bgcolor="FFFFFF">
	<td bgcolor="<%= adminColor("gray") %>" align="center">물류</td>	
	<td>
		<% 
		if companyrequest.results(0).physical = 0 then 
			response.write "물류시설 자체보유"
			response.write "("& companyrequest.results(0).physical_name & ")"
		else 
			response.write "물류전문업체 특정"
			response.write "("& companyrequest.results(0).physical_name & ")"
		end if
		%>
	</td>		
	<td bgcolor="<%= adminColor("gray") %>" align="center">제조</td>	
	<td>
		<% 
		if companyrequest.results(0).manufacturing = 0 then 
			response.write "생산공장 자체보유"
			response.write "("& companyrequest.results(0).manufacturing_name & ")"
		else 
			response.write "외부업체 특정"
			response.write "("& companyrequest.results(0).manufacturing_name & ")"
		end if
		%>
	</td>				
</tr>
<tr bgcolor="FFFFFF">
	<td bgcolor="<%= adminColor("gray") %>" align="center">산업재산권 취득</td>	
	<td>
		<%= companyrequest.results(0).industrial %>
	</td>
		
	<td bgcolor="<%= adminColor("gray") %>" align="center">라이센스 취득</td>	
	<td>
		<%= companyrequest.results(0).license %>
	</td>				
</tr>	
<tr bgcolor="FFFFFF">
	<td bgcolor="<%= adminColor("gray") %>" align="center">유통방법</td>	
	<td>
		<% 
		if left(companyrequest.results(0).utong,1) <> 0 then response.write "아직미판매 "
		if mid(companyrequest.results(0).utong,3,1) <> 0 then response.write "백화점 "
		if mid(companyrequest.results(0).utong,5,1) <> 0 then response.write "할인점 "	 
		if mid(companyrequest.results(0).utong,7,1) <> 0 then response.write "대리점 "
		if mid(companyrequest.results(0).utong,9,1) <> 0 then response.write "종합쇼핑몰 "
		if mid(companyrequest.results(0).utong,11,1) <> 0 then response.write "홈쇼핑 "	
		if mid(companyrequest.results(0).utong,13,1) <> 0 then response.write "타사거래처 "
		if mid(companyrequest.results(0).utong,15,1) <> 0 then response.write "자사몰 "	
		if right(companyrequest.results(0).utong,1) <> 0 then response.write "자사숍 "				
		%>
	</td>
		
	<td bgcolor="<%= adminColor("gray") %>" align="center">사업자형태</td>	
	<td>
		<% 
		if companyrequest.results(0).tax = 0 then 
			response.write "간이 "
		elseif  companyrequest.results(0).tax = 1 then 
			response.write "면세 "
		elseif  companyrequest.results(0).tax = 2 then 
			response.write "일반 "			
		else
			response.write "법인 "
		end if
		%>
	</td>				
</tr>	
<tr bgcolor="FFFFFF">
	<td bgcolor="<%= adminColor("gray") %>" align="center">회사URL</td>	
	<td>
		<%
			arrUrl = split(companyrequest.results(0).companyurl,",")
			if ubound(arrUrl)>0 then
				Response.Write "<a href='"
				if Left(arrUrl(0),7) <> "http://" and Left(arrUrl(0),8) <> "https://" then Response.Write "http://"
				Response.Write arrUrl(0) & "' target='_blank'>" & arrUrl(0) & "</a>"
				Response.Write "<br><br><b>입점쇼핑몰</b> : " & arrUrl(1)
			else
				Response.Write "<a href='"
				if Left(companyrequest.results(0).companyurl,7) <> "http://" and Left(companyrequest.results(0).companyurl,8) <> "https://" then Response.Write "http://"
				Response.Write companyrequest.results(0).companyurl & "' target='_blank'>" & companyrequest.results(0).companyurl & "</a>"
			end if
		%>
	</td>
		
	<td bgcolor="<%= adminColor("gray") %>" align="center">구분</td>	
	<td>
		<%= companyrequest.code2name(companyrequest.results(0).reqcd) %>
	</td>				
</tr>
<tr bgcolor="FFFFFF">		
	<td bgcolor="<%= adminColor("gray") %>" align="center">상품명(브랜드명)</td>	
	<td colspan=3>
		<%= nl2br(db2html(companyrequest.results(0).reqcomment)) %>
	</td>
			
</tr>
<tr bgcolor="FFFFFF">		
	<td bgcolor="<%= adminColor("gray") %>" align="center">첨부파일</td>	
	<td>
		<% if (companyrequest.results(0).attachfile <> "") then %>
			<a href="//imgstatic.10x10.co.kr<%= companyrequest.results(0).attachfile %>" target="_blank">다운받기</a>
		<% else %>
			없음
		<% end if %>
	</td>
					
	<td bgcolor="<%= adminColor("gray") %>" align="center">처리상태</td>	
	<td>
		<% if (IsNull(companyrequest.results(0).finishdate) = true) then %>
			미완료
		<% else %>
			<%= FormatDate(companyrequest.results(0).finishdate, "0000-00-00") %>
		<% end if %>
	</td>
</tr>
<tr bgcolor="FFFFFF"> 		
	<td bgcolor="<%= adminColor("gray") %>" align="center">회사설명</td>	
	<td colspan=3>
		<%= nl2br(db2html(companyrequest.results(0).companycomments)) %>
	</td>				
</tr>
<tr bgcolor="FFFFFF"> 			
	<td colspan=4 align="left">
	<input type="button" value="프린트" class="button" onclick="javascript:window.print();">
	</td>				
</tr>
</table>
<!-- 업체정보 끝 -->

<br>

<!-- 정보 변경  시작-->
<table width="100%" cellpadding="5" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="FFFFFF">		
	<td colspan=3><b><font color="blue">업체정보 변경</font></b></td>							
</tr>
<tr bgcolor="FFFFFF">		
	<td width="100">카테고리 변경</td>	
	<td width="300"><%if not isNull(companyrequest.results(0).dispcate) then%>
		<span><%=companyrequest.results(0).dispcatename1%> > <%=companyrequest.results(0).dispcatename2%></span>
		<%end if%>
		<div style="color:gray"> <%if companyrequest.results(0).cd1<>"" then%>관리: <% Drawcatelarge "catelargebox",companyrequest.results(i).cd1 %>(<% Drawcatemid companyrequest.results(0).cd1,"catemidbox",companyrequest.results(0).cd2 %>)<%end if%></div>
	</td>					
	<td>
		<!-- #include virtual="/common/module/dispCateSelectBoxDepth.asp"-->  
		<input type=button value="변경" onclick="catesubmit();">			
	</td>	
</tr>
<tr bgcolor="FFFFFF">		
	<td>판매형태 변경</td>	
	<td>
		<% if companyrequest.results(0).sellgubun="Y" then %>
		ON-Line/OFF-Line
		<% elseif companyrequest.results(0).sellgubun="N" then%>
		ON-Line
		<% elseif companyrequest.results(0).sellgubun="F" then%>
		OFF-Line
		<% else %>
		기타
		<% end if %>
	</td>					
	<td>
		<select name="sellgubun" class="a">
			<option value="Y">ON-Line/OFF-Line</option>
			<option value="N">ON-Line</option>
			<option value="F">OFF-Line</option>
		</select>
		<input type=button value="변경" onclick="sellsubmit();">			
	</td>	
</tr>
<tr bgcolor="FFFFFF">		
	<td>입점여부</td>	
	<td>
		<% if companyrequest.results(i).ipjumYN="Y" then response.write "입점완료" %>
		<% if companyrequest.results(i).ipjumYN="N" then response.write "미입점" %>
	</td>				
	<td>
	<!--<select name="ipjumYN" class="a">
		<option value="Y">입점 완료</option>
		<option value="N">미 입점</option>
	</select>-->
	<!--<input type=button value="변경" onclick="ipjumYNsubmit();">-->
	</td>	
</tr>	
<tr bgcolor="FFFFFF">		
	<td colspan=3>
		<input type="button" value=" 완료처리 " onclick="SubmitForm()" class="button">&nbsp;&nbsp;
		<% if companyrequest.results(i).fisusing="Y" then %>
			<input type="button" value="삭제" onclick="delreq()" class="button">&nbsp;&nbsp;
		<% end if %>
		<input type="button" value=" 입점프로세스 보내기 " onclick="AddNewBrand()" class="button"> 
	</td>
</tr>
</table>
</form>
<!-- 정보 변경  끝-->

<!-- 코멘트 부분 -->
<form name="commfrm" method="post" action="" onsubmit="return false" style="margin:0px;">
<table width="100%" cellpadding="5" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="FFFFFF">
	<td colspan=3><b><font color="blue">업체에게 메일보내기</font></b></td>
</tr>

<% if commmode="" and companyrequest.results(0).replyuser <>"" then %>
	<tr bgcolor="FFFFFF">
		<td width="10%" valign="top">
		작성: <%= db2html(companyrequest.results(0).replyuser) %>
		</td>
		<td width="75%" valign="top">
		<%= nl2br(db2html(companyrequest.results(0).replycomment)) %>
		</td>
		<td width="15%">
		<input type="button" value="수정" onclick="javascript:editcomm();">
		</td>
	</tr>
	<tr bgcolor="FFFFFF" align="left">
		<td colspan=3><input type="button" value="mail보내기" onclick="javascript:sendmail();">	</td>
	</tr>

<% 
'//수정모드
elseif commmode="edit" then
%>
	<tr bgcolor="FFFFFF">
		<td width="10%" valign="top">
			작성: <%= session("ssBctCname") %>
		</td>
		<td valign="top">
			<textarea name="comment" rows=10 cols=95><%= db2html(companyrequest.results(0).replycomment) %></textarea>
		</td>
		<td>
			<input type="button" value="저장" onclick="javascript:savecomm();">
		</td>
	</tr>
	
<% 
'//작성모드
elseif companyrequest.results(0).replyuser ="" then
%>
	<tr bgcolor="FFFFFF">
		<td valign="top">
			작성: <%= session("ssBctCname") %>
		</td>
		<td valign="top">
			<textarea name="comment" rows=10 cols=95></textarea>
		</td>
		<td>
			<input type="button" value="저장" onclick="javascript:savecomm();">
		</td>
	</tr>
<% end if %>

</table>
</form>

<form name="frm" method="post" action="" onsubmit="return false" style="margin:0px;">
	<input type="hidden" name="id" value="<%= companyrequest.results(0).id %>">
	<input type="hidden" name="pg" value="<%= page %>">
	<input type="hidden" name="mode" value="">
	<input type="hidden" name="disp" value=""> 
	<input type="hidden" name="sellgubun" value="">
	<input type="hidden" name="ipjumYN" value="">
	<input type="hidden" name="research" value="on">
	<input type="hidden" name="gubun" value="<%= gubun%>" >
	<input type="hidden" name="onlymifinish" value="<%=onlymifinish%>">
	<input type="hidden" name="catevalue" value="<%=catevalue%>">
	<input type="hidden" name="searchkey" value="<%=searchkey%>">
	<input type="hidden" name="commmode" value="">
	<input type="hidden" name="user" value="">
	<input type="hidden" name="comment" value="">
</form>
<form name="frmmail" method="post" action="/admin/board/upche/req_mail.asp" onsubmit="return false" style="margin:0px;">
	<input type="hidden" name="user" value="<%= session("ssBctCname") %>">
	<input type="hidden" name="userid" value="<%= session("ssBctId") %>">
	<input type="hidden" name="mailname" value="<%= companyrequest.results(0).chargename %>">
	<input type="hidden" name="mailto" value="<%= companyrequest.results(0).email %>">
	<input type="hidden" name="content" value="<%= companyrequest.results(0).replycomment %>">
	<input type="hidden" name="id" value="<%= companyrequest.results(0).id %>">
</form>
 
<script type="text/javascript">

//대카테고리선택시 중카테고리 셋팅
function searchCD2(paramCodeLarge) {
		
	resetLeftCountrySelect() ;		
	resetLeftCitySelect() ;
	
	if(paramCodeLarge != '') {
		FrameSearchCategory.location.href="/admin/CategoryMaster/frame_category_select.asp?search_code=" + paramCodeLarge + "&form_name=f&element_name=cd2";
	}
}

//중카테고리 선택시 소카테고리 셋팅	
function searchCD3(paramCodeMid) {	
	resetLeftCitySelect() ;
	
	if(paramCodeMid != '') {
		FrameSearchCategory.location.href="/admin/CategoryMaster/frame_category_select.asp?search_code=" + paramCodeMid + "&form_name=f&element_name=cd3";
	}	 
}

//대카테고리 초기화
function resetLeftCountrySelect() {
	document.f.cd2.length = 1;
	document.f.cd2.selectedIndex = 0 ;
}

		
//중카테고리 초기화
function resetLeftCitySelect() {
	document.f.cd3.length = 1;
	document.f.cd3.selectedIndex = 0 ;
}

</script>

<%
set companyrequest=nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
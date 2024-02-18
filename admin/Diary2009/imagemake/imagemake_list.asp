<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 다이어리스토리
' History : 2008.10.12 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/diary2009/classes/DiaryCls.asp"-->

<%
dim research,isusing, fixtype, linktype, poscode, validdate
dim page

	isusing = request("isusing")
	research= request("research")
	poscode = request("poscode")
	fixtype = request("fixtype")
	page    = request("page")
	validdate= request("validdate")

if trim(research)<>"on" and isusing="" then
    isusing = "Y"
    validdate = "on"
elseif trim(research)="on" and isusing="" then
	isusing = ""
end if
'response.write isusing &"//"& research
'response.end
if page="" then page=1

dim oposcode
set oposcode = new DiaryCls
	oposcode.FRectPosCode = poscode
	if (poscode<>"") then
	    oposcode.fposcode_oneitem
	end if

dim oMainContents
set oMainContents = new DiaryCls
	oMainContents.FPageSize = 20
	oMainContents.FCurrPage = page
	oMainContents.FRectIsusing = isusing
	oMainContents.FRectPosCode = poscode
	oMainContents.FRectvaliddate = validdate
	oMainContents.fcontents_list

dim i

%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" lang="ko" xml:lang="ko">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<link rel="stylesheet" type="text/css" href="/css/adminDefault.css" />
<link rel="stylesheet" type="text/css" href="/css/adminCommon.css" />
<script type="text/javascript">
function AnSelectAllFrame(bool){
	var frm;
	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			if (frm.cksel.disabled!=true){
				frm.cksel.checked = bool;
				AnCheckClick(frm.cksel);
			}
		}
	}
}

function AnCheckClick(e){
	if (e.checked)
		hL(e);
	else
		dL(e);
}

function ckAll(icomp){
	var bool = icomp.checked;
	AnSelectAllFrame(bool);
}

function CheckSelected(){
	var pass=false;
	var frm;

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			pass = ((pass)||(frm.cksel.checked));
		}
	}

	if (!pass) {
		return false;
	}
	return true;
}

// 메인 플래쉬 실서버 적용
function AssignFlashReal(upfrm,poscode,imagecount){
if (!CheckSelected()){
		alert('선택아이템이 없습니다.');
		return;
	}
	var frm;
		for (var i=0;i<document.forms.length;i++){
			frm = document.forms[i];
			if (frm.name.substr(0,9)=="frmBuyPrc") {
				if (frm.cksel.checked){
					upfrm.fidx.value = upfrm.fidx.value + frm.idx.value + "," ;

				}
			}
		}
			var tot;
			tot = upfrm.fidx.value;
			upfrm.fidx.value = ""
		var AssignFlashReal;
		AssignFlashReal = window.open("<%=wwwUrl%>/chtml/diary/diary_"+poscode+".asp?idx=" +tot + '&poscode='+poscode+'&imagecount='+imagecount, "AssignFlashReal","width=800,height=600,scrollbars=yes,resizable=yes");
		AssignFlashReal.focus();
}

// 엠디픽플래쉬 실서버 적용
function AssignFlashReal_mdpick(upfrm,poscode,imagecount){
if (!CheckSelected()){
		alert('선택아이템이 없습니다.');
		return;
	}
	var frm;
		for (var i=0;i<document.forms.length;i++){
			frm = document.forms[i];
			if (frm.name.substr(0,9)=="frmBuyPrc") {
				if (frm.cksel.checked){
					upfrm.fidx.value = upfrm.fidx.value + frm.idx.value + "," ;

				}
			}
		}
			var tot;
			tot = upfrm.fidx.value;
			upfrm.fidx.value = ""
		var AssignFlashReal;
		AssignFlashReal = window.open("<%=wwwUrl%>/chtml/diary_mdpick_flashmake.asp?idx=" +tot + '&poscode='+poscode+'&imagecount='+imagecount, "AssignFlashReal","width=800,height=600,scrollbars=yes,resizable=yes");
		AssignFlashReal.focus();
}

//포스 코드 등록 & 수정
function popPosCodeManage(){
    var popPosCodeManage = window.open('/admin/diary2009/imagemake/imagemake_poscode.asp','popPosCodeManage','width=1024,height=768,scrollbars=yes,resizable=yes');
    popPosCodeManage.focus();
}

//이미지신규등록 & 수정
function AddNewMainContents(idx){
    var AddNewMainContents = window.open('/admin/diary2009/imagemake/imagemake_contents.asp?idx='+ idx,'AddNewMainContents','width=1024,height=768,scrollbars=yes,resizable=yes');
    AddNewMainContents.focus();
}

function AssignTest(){
    if (document.frm.poscode.value == ""){
		alert("적용위치를 선택해주세요");
		document.frm.poscode.focus();
	}
	else{
		 var popwin = window.open('','refreshFrm_Main_Test','');
		 popwin.focus();
		 refreshFrm.target = "refreshFrm_Main_Test";
		 refreshFrm.action = "<%=uploadUrl%>/chtml/othermall_contents_Test_JS.asp?poscode=" + document.frm.poscode.value;
		 refreshFrm.submit();
	}
}

function AssignReal(){
    if (document.frm.poscode.value == ""){
		alert("적용위치를 선택해주세요");
		document.frm.poscode.focus();
	}
	else{
		 var popwin = window.open('','refreshFrm_Main','');
		 popwin.focus();
		 refreshFrm.target = "refreshFrm_Main";
		 refreshFrm.action = "<%=uploadUrl%>/chtml/othermall_make_main_contents_JS.asp?poscode=" + document.frm.poscode.value;
		 refreshFrm.submit();
	}
}


function AssignDailyTest(idx){
	 var popwin = window.open('','refreshFrm_Main_Test','');
	 popwin.focus();
	 refreshFrm.target = "refreshFrm_Main_Test";
	 refreshFrm.action = "<%=othermall%>/chtml/othermall_make_main_contents_byidx_Test_JS.asp?idx=" + idx;
	 refreshFrm.submit();
}

function AssignDailyReal(idx,poscode,imagecount){
	<% If poscode = "14" Then %>
    var AddNewMainContents = window.open('<%=wwwUrl%>/chtml/mobile/diary.asp?idx='+ idx + '&poscode='+poscode+'&imagecount='+imagecount,'AddNewMainContents','width=800,height=600,scrollbars=yes,resizable=yes');
    <% Else %>
	var AddNewMainContents = window.open('<%=wwwUrl%>/chtml/diary/diary.asp?idx='+ idx + '&poscode='+poscode+'&imagecount='+imagecount,'AddNewMainContents','width=800,height=600,scrollbars=yes,resizable=yes');
	<% End If %>
    AddNewMainContents.focus();
}

function AssignXMLReal(){
    var AddNewXMainContents = window.open('<%=wwwUrl%>/chtml/diary/diary_xml.asp','AddNewXMainContents','width=800,height=600,scrollbars=yes,resizable=yes');
    AddNewXMainContents.focus();
}
</script>
</head>
<body>
<div class="contSectFix scrl">
	<!-- 검색 시작 -->
	<div class="pad20">
		<table class="tbType1 listTb">
			<form name="frm" method="get" action="">
			<input type="hidden" name="page" value="1">
			<input type="hidden" name="research" value="on">
			<input type="hidden" name="menupos" value="<%= menupos %>">
			<input type="hidden" name="fidx">
			<tr bgcolor="#FFFFFF">
				<td width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
				<td style="text-align:left;">
					<!--<input type="checkbox" name="validdate" <% if validdate="on" then response.write "checked" %> >종료이전-->
					사용구분
					<select name="isusing">
					<option value="">전체
					<option value="Y" <% if isusing="Y" then response.write "selected" %> >사용함
					<option value="N" <% if isusing="N" then response.write "selected" %> >사용안함
					</select>
					적용구분
					<% call DrawMainPosCodeCombo("poscode", poscode,"") %>
				</td>
				<td width="50" bgcolor="<%= adminColor("gray") %>">
					<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
				</td>
			</tr>
			</form>
		</table>
		<div class="tPad15">
			<table class="tbType1 listTb">
				<tr>
					<td style="text-align:left;">※적용구분을 선택하셔야 실서버 반영 버튼이 생김니다.
						<% if (poscode<>"") then %>
							<% if oposcode.FOneItem.fimagetype="flash" then %>
								<% if oposcode.FOneItem.fposcode =2 then %>
								<a href="javascript:AssignFlashReal(frm,<%= poscode %>,<%=oposcode.FOneItem.fimagecount%>);"><img src="/images/refreshcpage.gif" border="0"> Flash Real 적용</a>
								<% else %>
								<a href="javascript:AssignFlashReal(frm,<%= poscode %>,<%=oposcode.FOneItem.fimagecount%>);"><img src="/images/refreshcpage.gif" border="0"> Flash Real 적용</a>
								<% end if %>
							<% elseif oposcode.FOneItem.fimagetype="multi" then %>
								<a href="javascript:AssignTest('<%= poscode %>');"><img src="/images/icon_search.jpg" border="0"> 미리보기</a>
								&nbsp;&nbsp;
								<a href="javascript:AssignReal('<%= poscode %>');"><img src="/images/refreshcpage.gif" border="0"> Real 적용</a>
							<% end if %>
						<% end if %>
						<% if C_ADMIN_AUTH then %>
						<input type="button" value="코드관리" class="button" onClick="popPosCodeManage();">
						<% end if %>
						<input type="button" value="신규등록" class="button" onClick="AddNewMainContents('0');">
					</td>
				</tr>
			</table>
		</div>		
		<div class="tPad15">
			<table class="tbType1 listTb">
			<% if oMainContents.FResultCount > 0 then %>
				<tr height="25" bgcolor="FFFFFF">
					<td colspan="15" style="text-align:left;">
						검색결과 : <b><%= oMainContents.FTotalCount %></b>
						&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
						<% If CStr(poscode) = "1" Or CStr(poscode) = "11" Then %>※ <font color="red"><b>메인 멀티 적용 : 우선 순위가 높은 순서부터 3개가 적용이 됩니다.</b></font><a href="javascript:AssignXMLReal('0','4','5');"><img src="/images/refreshcpage.gif" border="0"> Real 적용</a><% End If %>
					</td>
				</tr>
				<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
					<!--<td align="center"><input type="checkbox" name="ckall" onclick="ckAll(this)"></td>//-->
					<td align="center">Idx</td>
					<td align="center">Image(수정)</td>
					<td align="center">구분명</td>
					<td align="center">상품코드</td>
					<td align="center">LinkType</td>
					<td align="center">우선순위</td>
					<td align="center">사용여부</td>
					<td align="center">기간</td>
					<td align="center">등록일</td>
						<% if (poscode<>"") then %>
							<% if Not(oMainContents.FItemList(i).FIsusing="N" or oMainContents.FItemList(i).Fimagetype="flash") then %>
								<td></td>
							<% end if %>
						<% end if %>
				</tr>
				<tr align="center" bgcolor="#FFFFFF">
				<% for i=0 to oMainContents.FResultCount - 1 %>
				<form action="" name="frmBuyPrc<%=i%>" method="get">			<!--for문 안에서 i 값을 가지고 루프-->
					<% if oMainContents.FItemList(i).FIsusing="N" or left(oMainContents.FItemList(i).fevent_end,10) < date() then %>
						<tr bgcolor="#DDDDDD">
					<% else %>
						<tr bgcolor="#FFFFFF">
					<% end if %>

					<!--<td align="center"><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"></td>//-->
					<td align="center"><%= oMainContents.FItemList(i).Fidx %><input type="hidden" name="idx" value="<%= oMainContents.FItemList(i).Fidx %>"></td>
					<td align="center">
					<% if poscode="200" then %>
						<a href="javascript:AddNewMainContents('<%= oMainContents.FItemList(i).Fidx %>');">
						<img width=40 height=40 src="<%= oMainContents.FItemList(i).fimagesmall %>" border="0">
						</a>
					<% else %>
						<a href="javascript:AddNewMainContents('<%= oMainContents.FItemList(i).Fidx %>');">
						<img width=40 height=40 src="<%=uploadUrl%>/diary/main/<%= oMainContents.FItemList(i).fimagepath %>" border="0">
						</a>
					<% end if %>
					</td>
					<td align="center"><a href="?poscode=<%= oMainContents.FItemList(i).Fposcode %>"><%= oMainContents.FItemList(i).Fposname %></a></td>
					<td align="center"><%= oMainContents.FItemList(i).fevt_code %></td>
					<td align="center"><%= oMainContents.FItemList(i).fimagetype %></td>
					<td align="center"><%= oMainContents.FItemList(i).fimage_order %></td>
					<td align="center"><%= oMainContents.FItemList(i).FIsusing %></td>
					<td align="center"><%= left(oMainContents.FItemList(i).fevent_start,10) %> ~ <%= left(oMainContents.FItemList(i).fevent_end,10) %>
					<% if Not isNull(oMainContents.FItemList(i).Fusedate) Then Response.Write "<br>그룹코드 : " & oMainContents.FItemList(i).Fusedate End If %>
					</td>
					<td align="center"><%= oMainContents.FItemList(i).fregdate %></td>

					<% if (poscode<>"") and poscode<>"402" then %>
						<% if Not(oMainContents.FItemList(i).FIsusing="N" or oMainContents.FItemList(i).Fimagetype="flash" or oMainContents.FItemList(i).Fimagetype="multi") then %>
							<td>
								<!--<a href="javascript:AssignDailyTest('<%= oMainContents.FItemList(i).Fidx %>');"><img src="/images/icon_search.jpg" border="0"> 미리보기</a>
								&nbsp;//-->
								<% If poscode <> 1 and poscode <> 11  Then %>
									<% If poscode<>16 and poscode<> 17 Then %>
										<a href="javascript:AssignDailyReal('<%= oMainContents.FItemList(i).Fidx %>','<%= poscode %>','<%=oMainContents.FItemList(i).fimagecount%>');"><img src="/images/refreshcpage.gif" border="0"> Real 적용</a>
									<% end if %>
								<% end if %>
							</td>
						<% end if %>
					<% end if %>
				</tr>
				</form>
				<% next %>
				</tr>
				<% else %>
				<tr bgcolor="#FFFFFF">
					<td colspan="7" align="center" class="page_link">[검색결과가 없습니다.]</td>
				</tr>
				<% end if %>
				<tr height="25" bgcolor="FFFFFF">
					<td colspan="15" align="center">
						<% if oMainContents.HasPreScroll then %>
							<span class="list_link"><a href="?page=<%= oMainContents.StartScrollPage-1 %>&poscode=<%=poscode%>&isusing=<%=isusing%>&research=<%=research%>">[pre]</a></span>
						<% else %>
						[pre]
						<% end if %>
						<% for i = 0 + oMainContents.StartScrollPage to oMainContents.StartScrollPage + oMainContents.FScrollCount - 1 %>
							<% if (i > oMainContents.FTotalpage) then Exit for %>
							<% if CStr(i) = CStr(oMainContents.FCurrPage) then %>
							<span class="page_link"><font color="red"><b><%= i %></b></font></span>
							<% else %>
							<a href="?page=<%= i %>&poscode=<%=poscode%>&isusing=<%=isusing%>&research=<%=research%>" class="list_link"><font color="#000000"><%= i %></font></a>
							<% end if %>
						<% next %>
						<% if oMainContents.HasNextScroll then %>
							<span class="list_link"><a href="?page=<%= i %>&poscode=<%=poscode%>&isusing=<%=isusing%>&research=<%=research%>">[next]</a></span>
						<% else %>
						[next]
						<% end if %>
					</td>
				</tr>
			</table>
		</div>
	</div>
</div>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->


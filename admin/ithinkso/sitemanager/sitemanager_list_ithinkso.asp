<%@  codepage="65001" language="VBScript" %>
<% option explicit %>
<% response.Charset="UTF-8" %>
<% session.codePage = 65001 %>
<%
'###########################################################
' Description : 아이띵소 사이트 관리
' Hieditor : 2013.05.15 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
<script language="JavaScript" src="/js/xl.js"></script>
<script language="JavaScript" src="/js/common.js"></script>
<script language="JavaScript" src="/js/report.js"></script>
<script language="JavaScript" src="/cscenter/js/cscenter.js"></script>
<script language="JavaScript" src="/js/calendar.js"></script>

<link rel="stylesheet" href="/css/scm.css" type="text/css">
<script language='javascript'>

function PopMenuHelp(menupos){
	var popwin = window.open('/designer/menu/help.asp?menupos=' + menupos,'PopMenuHelp_a','width=900, height=600, scrollbars=yes,resizable=yes');
	popwin.focus();
}

function PopMenuEdit(menupos){
	var popwin = window.open('/admin/menu/pop_menu_edit.asp?mid=' + menupos,'PopMenuEdit','width=600, height=400, scrollbars=yes,resizable=yes');
	popwin.focus();
}

</script>

<% if session("sslgnMethod")<>"S" then %>
	<!-- USB키 처리 시작 (2008.06.23;허진원) -->
	<OBJECT ID='MaGerAuth' WIDTH='' HEIGHT=''	CLASSID='CLSID:781E60AE-A0AD-4A0D-A6A1-C9C060736CFC' codebase='/lib/util/MaGer/MagerAuth.cab#Version=1,0,2,4'></OBJECT>
	<script language="javascript" src="/js/check_USBToken.js"></script>
	<!-- USB키 처리 끝 -->
<% end if %>
</head>
<body bgcolor="#F4F4F4" onload="checkUSBKey()">

<!-- #include virtual="/lib/classes/ithinkso/sitemanager/sitemanager_cls_ithinkso_utf8.asp"-->

<%
dim research,isusing, fixtype, linktype, code, page, menupos, i
	isusing = request("isusing")
	research= request("research")
	code = request("code")
	fixtype = request("fixtype")
	page    = request("page")
	menupos = request("menupos")
	
if (research="") and (isusing="") then 
    isusing = "Y"
end if

if page="" then page=1

dim ocode
set ocode = new csitemanager_list
	ocode.frectcode = code
	ocode.frectisusing = "Y"
	
	if (code<>"") then
	    ocode.fsitemanager_code_one()
	end if

dim oContents
set oContents = new csitemanager_list
	oContents.FPageSize = 50
	oContents.FCurrPage = page
	oContents.FRectIsusing = isusing
	oContents.frectcode = code
	oContents.fsitemanager_list()
%>

<script language="javascript">

// 이미지 실서버 적용
function AssignmaindownbarnerReal(upfrm,code,imagecount){
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
		var AssignmaindownbarnerReal;
		if(code == "500")
		{
			AssignmaindownbarnerReal = window.open("<%=wwwithinksoweb%>/chtml/maindownbarner_make.asp?idx=" +tot + '&code='+code+'&imagecount='+imagecount, "AssignmaindownbarnerReal","width=400,height=300,scrollbars=yes,resizable=yes");
		}
		else if(code == "504")
		{
			AssignmaindownbarnerReal = window.open("<%=wwwithinksoweb%>/chtml/maindownbarner_make_2011.asp?idx=" +tot + '&code='+code+'&imagecount='+imagecount, "AssignmaindownbarnerReal","width=400,height=300,scrollbars=yes,resizable=yes");
		}
		AssignmaindownbarnerReal.focus();
}

function AssignimageReal(upfrm,code,imagecount){
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
		var AssignimageReal;
		AssignimageReal = window.open("<%=wwwithinksoweb%>/chtml/imagemake.asp?idx=" +tot + '&code='+code+'&imagecount='+imagecount, "AssignimageReal","width=400,height=300,scrollbars=yes,resizable=yes");
		AssignimageReal.focus();
}

function AssignXmlReal(upfrm,code,imagecount){
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
		var AssignXmlReal;
		AssignXmlReal = window.open("<%=wwwithinksoweb%>/chtml/xmlmake.asp?idx=" +tot + '&code='+code+'&imagecount='+imagecount, "AssignXmlReal","width=400,height=300,scrollbars=yes,resizable=yes");
		AssignXmlReal.focus();
}

//포스 코드 등록 & 수정
function popsitemanager_code(){
    var popsitemanager_code = window.open('/admin/ithinkso/sitemanager/sitemanager_code_ithinkso.asp','popsitemanager_code','width=1024,height=768,scrollbars=yes,resizable=yes');
    popsitemanager_code.focus();
}

//이미지신규등록 & 수정
function AddNewMainContents(idx){
    var AddNewMainContents = window.open('/admin/ithinkso/sitemanager/sitemanager_contents_ithinkso.asp?idx='+ idx,'AddNewMainContents','width=1024,height=768,scrollbars=yes,resizable=yes');
    AddNewMainContents.focus();
}

function frmsubmit(page){
	frm.page.value=page;
	frm.submit();
}

document.domain ='10x10.co.kr';

</script>

<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td>
		<table width="100%" align="center" cellpadding="1" cellspacing="0" class="a" bgcolor="#999999">
			<tr>
				<td width="400" style="padding:5; border-top:1px solid #999999;border-left:1px solid #999999;border-right:1px solid #999999"  background="/images/menubar_1px.gif">
					<font color="#333333"><b>[3PL]아이띵소사이트관리&gt;&gt;아이띵소해외사이트관리</b></font>
				</td>
				
				<td align="right" style="border-bottom:1px solid #999999" bgcolor="#F4F4F4">
					<!-- 마스터이상 메뉴권한 설정
					
					<a href="Javascript:PopMenuEdit('1491');"><img src="/images/icon_chgauth.gif" border="0" valign="bottom"></a> -->
					
					<!-- Help 설정
					
					<a href="Javascript:PopMenuHelp('1491');"><img src="/images/icon_help.gif" border="0" valign="bottom"></a> -->
				</td>
				
			</tr>
		</table>
	</td>
</tr>
</table>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="page" value="">
<input type="hidden" name="research" value="on">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="fidx">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
	    * 사용여부 : <% drawSelectBoxisusingYN "isusing", isusing, " onchange='frmsubmit("""")'" %>
		&nbsp;&nbsp;
		* 적용구분
		<% call DrawsitemanagerCode("code", code, " onchange='frmsubmit("""")'") %>
	</td>
	
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="frmsubmit('');">
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
	    <% 
	    '//적용구분 선택시에만 뿌림
	    if (code<>"") then 
	    	if ocode.ftotalcount >0 then
	    %>
			    <% if ocode.FOneItem.fimagetype="link" then %>
			    	<%
			    	if ocode.FOneItem.fcode = "100" then
			    	%>
						<a href="javascript:AssignimageReal(frm,<%= code %>,<%=ocode.FOneItem.fimagecount%>);"><img src="/images/refreshcpage.gif" border="0"> Real 적용</a>			    	
			    	<%
			    	end if	
			    	%>
			    <% elseif ocode.FOneItem.fimagetype="xml" then %>
					<!--<a href="javascript:AssignXmlReal(frm,<%'= code %>,<%'=ocode.FOneItem.fimagecount%>);"><img src="/images/refreshcpage.gif" border="0"> XML Real 적용</a>-->
			    <%
			    end if 
			    %>
			<% end if %>
	    <% end if %>
	    
	    ※ 적용구분을 검색하신후에 데이터를 등록 하시고, <font color="red">Real적용</font> 버튼을 누르셔야 프론트에 노출 됩니다.
	</td>
	<td align="right">
		<% if C_ADMIN_AUTH then %>
			<input type="button" value="코드관리" class="button" onClick="popsitemanager_code();">
		<% end if %>
		&nbsp;&nbsp;
		<input type="button" value="신규등록" class="button" onClick="javascript:AddNewMainContents('');">						
	</td>
	</tr>
</table>
<!-- 액션 끝 -->

<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		검색결과 : <b><%= oContents.FTotalCount %></b>
		&nbsp;
		페이지 : <b><%= page %>/ <%= oContents.FTotalPage %></b>	
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td align="center"><input type="checkbox" name="ckall" onclick="ckAll(this)"></td>
    <td align="center">Idx</td>
    <td align="center">Image</td>
    <td align="center">구분명</td>
    <td align="center">LinkType</td>
    <td align="center">우선순위</td>
    <td align="center">사용여부</td>
    <td align="center">등록일</td>
    <td align="center">비고</td>
</tr>
<% if oContents.FResultCount > 0 then %> 
<tr align="center" bgcolor="#FFFFFF">
<% for i=0 to oContents.FResultCount - 1 %>
<form action="" name="frmBuyPrc<%=i%>" method="get">			 		
	<% if oContents.FItemList(i).FIsusing="N" then %>
		<tr bgcolor="#DDDDDD" align="center">
	<% else %>
		<tr bgcolor="#FFFFFF" align="center">
	<% end if %>
	
	<td><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"></td>		
    <td><%= oContents.FItemList(i).Fidx %><input type="hidden" name="idx" value="<%= oContents.FItemList(i).Fidx %>"></td>
    <td>
    	<% if oContents.FItemList(i).fimagepath <> "" then %>
    		<img width=50 height=50 src="<%=uploadUrl%>/ithinkso/sitemanager/<%= oContents.FItemList(i).fimagepath %>" border="0">
    	<% end if %>
    </td>
    <td>
    	<%= oContents.FItemList(i).Fcodename %>
    </td>
    <td><%= oContents.FItemList(i).fimagetype %></td>
    <td><%= oContents.FItemList(i).fimage_order %></td>
    <td><%= oContents.FItemList(i).FIsusing %></td>
    <td><%= oContents.FItemList(i).fregdate %></td>
    <td><input type="button" value="수정" onclick="AddNewMainContents('<%= oContents.FItemList(i).Fidx %>');" class="button"></td>
</tr>
</form>	
<% next %>
</tr>   

<tr height="25" bgcolor="FFFFFF">
	<td colspan="15" align="center">
       	<% if oContents.HasPreScroll then %>
			<span class="list_link"><a href="javascript:frmsubmit('<%= oContents.StartScrollPage-1 %>');">[pre]</a></span>
		<% else %>
		[pre]
		<% end if %>
		<% for i = 0 + oContents.StartScrollPage to oContents.StartScrollPage + oContents.FScrollCount - 1 %>
			<% if (i > oContents.FTotalpage) then Exit for %>
			<% if CStr(i) = CStr(oContents.FCurrPage) then %>
			<span class="page_link"><font color="red"><b><%= i %></b></font></span>
			<% else %>
			<a href="javascript:frmsubmit('<%= i %>');" class="list_link"><font color="#000000"><%= i %></font></a>
			<% end if %>
		<% next %>
		<% if oContents.HasNextScroll then %>
			<span class="list_link"><a href="javascript:frmsubmit('<%= i %>');">[next]</a></span>
		<% else %>
		[next]
		<% end if %>
	</td>
</tr>

<% else %>
<tr bgcolor="#FFFFFF">
	<td colspan="15" align="center" class="page_link">[검색결과가 없습니다.]</td>
</tr>
<% end if %>
</table>

<%
session.codePage = 949

set ocode = nothing
set oContents = nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
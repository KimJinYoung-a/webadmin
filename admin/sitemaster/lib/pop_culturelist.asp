<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : Culture Station Event
' History : 2009.04.02 한용민 생성
'           2012.01.12 허진원; 모바일 추가, 폼방식 수정
'           2013.06.04 허진원; 진행상태에 따른 배경색 추가
'			2018.09.27 정태훈 : 컬쳐스테이션 이벤트DB 이전 변경 적용
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/eventmanage/common/event_function.asp"-->
<!-- #include virtual="/lib/classes/culturestation/culturestation_class.asp"-->

<%
Dim oip,i,page,evt_type_search,isusing_search,evt_code_search, evt_name_search, evt_code_count, evt_mobile_yn, evt_partner_search
Dim edid, emid, rowColor
Dim sDate,sSdate,sEdate, sortMtd, srchStat

Dim gubun , pcode , pidx

	evt_code_search = request("evt_code_search")
	evt_name_search = request("evt_name_search")
	evt_partner_search = request("evt_partner_search")
	evt_type_search = request("evt_type_searchbox")
	isusing_search = request("isusing_searchbox")
	evt_code_count = request("evt_code_countbox")
	evt_mobile_yn = request("evt_mobile_yn")
	menupos = request("menupos")
	page = request("page")
	sortMtd = request("sortMtd")
	srchStat = request("srchStat")

	gubun = request("gubun")
	pcode = request("poscode")
	pidx = request("pidx")

	edid  		= requestCheckVar(Request("selDId"),32)		'담당 디자이너
	emid  		= requestCheckVar(Request("selMId"),32)		'담당 MD

	sDate 		= requestCheckVar(Request("selDate"),1)  	'기간
	sSdate 		= requestCheckVar(Request("iSD"),10)
	sEdate 		= requestCheckVar(Request("iED"),10)

	if page = "" then page = 1

'// 이벤트 리스트
set oip = new cevent_list
	oip.FPageSize = 50
	oip.FCurrPage = page
	oip.frectevt_type = evt_type_search
	oip.frectisusing = isusing_search
	oip.frectevt_code = evt_code_search
	oip.frectevt_partner = evt_partner_search
	oip.frectevt_name = evt_name_search
	oip.frectevt_code_count = evt_code_count
	oip.frectSortMethod = sortMtd
	oip.frectStatus = srchStat

	oip.fedid	= edid
	oip.femid	= emid

	oip.fdate	= sDate
	oip.fsdate	= sSdate
	oip.fedate	= sEdate

	oip.GetCulturePopSelectList()
%>

<script language="javascript">

function event_edit(evt_code){
	var event_edit = window.open('/admin/culturestation/event_edit.asp?evt_code='+evt_code,'addreg','width=800,height=768,scrollbars=yes,resizable=yes');
	event_edit.focus();
}

function AnSelectAllFrame(bool){
	var frm = document.frmBuyPrc;
	if(frm.chkitem.length>1) {
		for (var i=0;i<frm.chkitem.length;i++){
			if (frm.chkitem[i].disabled!=true){
				frm.chkitem[i].checked = bool;
				AnCheckClick(frm.chkitem[i]);
			}
		}
	} else {
		frm.chkitem.checked = bool;
		AnCheckClick(frm.chkitem);
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

	var frm = document.frmBuyPrc;
	if(frm.chkitem.length>1) {
		for (var i=0;i<frm.chkitem.length;i++){
			pass = ((pass)||(frm.chkitem[i].checked));
		}
	} else {
		pass = ((pass)||(frm.chkitem.checked));
	}

	if (!pass) {
		return false;
	}
	return true;
}

function comment_list(evt_code){

	 var comment_list = window.open('/admin/culturestation/event_comment_list.asp?evt_code='+evt_code,'comment_list','width=800,height=600,scrollbars=yes,resizable=yes');
	 comment_list.focus();

}

function goPage(pg) {
	var frm = document.frm;
	frm.evt_code.value="";
	frm.page.value=pg;
	frm.action="";
	frm.submit();
}
function RefreshMainCorItemRec(upfrm){
	if (!CheckSelected()){
		alert('선택아이템이 없습니다.');
		return;
	}

	var frm = document.frmBuyPrc;
	if(frm.chkitem.length>1) {
		for (var i=0;i<frm.chkitem.length;i++){
			if (frm.chkitem[i].checked)
				upfrm.evt_code.value = upfrm.evt_code.value + frm.evt_code[i].value + "," ;
		}
	} else {
		if (frm.chkitem.checked)
			upfrm.evt_code.value = upfrm.evt_code.value + frm.evt_code.value + "," ;
	}

	var tot;
	tot = upfrm.evt_code.value;
	upfrm.evt_code.value = ""
	var AssignReal;
	AssignReal = window.open("<%=wwwUrl%>/chtml/main_curture_make12banner.asp?evt_code=" +tot, "AssignReal","width=400,height=300,scrollbars=yes,resizable=yes");
	AssignReal.focus();
}


// 부모창에 값 넘기기
function jsSetEvtCont(ieC){
	if(typeof(opener.document) == "object"){
		opener.location.href = "popmaincontentsedit.asp?eC="+ieC+"&gubun=<%=gubun%>&poscode=<%=pcode%>&idx=<%=pidx%>";
		window.close();
	}
}

// 부모창에 값 넘기기
function jsSetEvtContMobile(ieC){
	if(typeof(opener.document) == "object"){
		opener.location.href = "/admin/mobile/popmaincontentsedit.asp?eC="+ieC+"&poscode=<%=pcode%>&idx=<%=pidx%>";
		window.close();
	}
}

function TnSearchEvtSelect(objval){
	if(objval=="evt_code_search"){
		$("#evt_code_search").css("display","");
		$("#evt_name_search").css("display","none");
		$("#evt_partner_search").css("display","none");
	}else if(objval=="evt_name_search"){
		$("#evt_code_search").css("display","none");
		$("#evt_name_search").css("display","");
		$("#evt_partner_search").css("display","none");
	}else{
		$("#evt_code_search").css("display","none");
		$("#evt_name_search").css("display","none");
		$("#evt_partner_search").css("display","");
	}
}
</script>
<script type="text/javascript" src="/js/jquery-2.2.2.min.js"></script>
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method=get action="">
	<input type="hidden" name="evt_code">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="page">
	<input type="hidden" name="sortMtd" value="<%=sortMtd%>">
	<input type="hidden" name="poscode" value="<%=pcode%>">
	<input type="hidden" name="gubun" value="<%=gubun%>">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
		<td align="left">기간 :
			<select name="selDate" class="select">
		    	<option value="S" <%if Cstr(sDate) = "S" THEN %>selected<%END IF%>>시작일 기준</option>
		    	<option value="E" <%if Cstr(sDate) = "E" THEN %>selected<%END IF%>>종료일 기준</option>
		    	<option value="V" <%if Cstr(sDate) = "V" THEN %>selected<%END IF%>>발표일 기준</option>
			</select>
	        <input id="iSD" name="iSD" value="<%=sSdate%>" class="text" size="10" maxlength="10" /><img src="http://scm.10x10.co.kr/images/calicon.gif" id="iSD_trigger" border="0" style="cursor:pointer" align="absmiddle" /> ~
	        <input id="iED" name="iED" value="<%=sEdate%>" class="text" size="10" maxlength="10" /><img src="http://scm.10x10.co.kr/images/calicon.gif" id="iED_trigger" border="0" style="cursor:pointer" align="absmiddle" />
	        /
	        <input type="checkbox" name="srchStat" value="Y" <%=chkIIF(srchStat="Y","checked","")%> />진행중인 이벤트만 보기
			<script language="javascript">
				var CAL_Start = new Calendar({
					inputField : "iSD", trigger    : "iSD_trigger",
					onSelect: function() {
						var date = Calendar.intToDate(this.selection.get());
						CAL_End.args.min = date;
						CAL_End.redraw();
						this.hide();
					}, bottomBar: true, dateFormat: "%Y-%m-%d"
				});
				var CAL_End = new Calendar({
					inputField : "iED", trigger    : "iED_trigger",
					onSelect: function() {
						var date = Calendar.intToDate(this.selection.get());
						CAL_Start.args.max = date;
						CAL_Start.redraw();
						this.hide();
					}, bottomBar: true, dateFormat: "%Y-%m-%d"
				});
			</script>
		</td>
		<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="검색" onClick="frm.page.value=1;frm.submit();">
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">
			구분:
			<select name="evt_type_searchbox" value="<%=evt_type_search%>" class="select">
				<option value="" <% if evt_type_search = "" then response.write " selected" %>>전체</option>
				<option value="0" <% if evt_type_search = "0" then response.write " selected" %>>느껴봐</option>
				<option value="1" <% if evt_type_search = "1" then response.write " selected" %>>읽어봐</option>
				<option value="2" <% if evt_type_search = "2" then response.write " selected" %>>들어봐</option>
			</select> /
			사용여부:
			<select name="isusing_searchbox" value="<%=isusing_search%>" class="select">
				<option value="" <% if isusing_search = "" then response.write " selected" %>>전체</option>
				<option value="Y" <% if isusing_search = "Y" then response.write " selected" %>>Y</option>
				<option value="N" <% if isusing_search = "N" then response.write " selected" %>>N</option>
			</select> /
			코멘트 사용:
			<select name="evt_code_countbox" value="<%=evt_code_count%>" class="select">
				<option value="" <% if evt_code_count = "" then response.write " selected" %>>전체</option>
				<option value="Y" <% if evt_code_count = "Y" then response.write " selected" %>>Y</option>
				<option value="N" <% if evt_code_count = "N" then response.write " selected" %>>N</option>
			</select> /
			WD담당: <%sbGetDesignerid "selDId",edid, "onChange='javascript:document.frm.submit();'"%>
			마케팅담당: <%sbGetMKTid "selMId",emid, "onChange='javascript:document.frm.submit();'"%>
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >

		<td align="left">
			검색어 : 	<select name="ses" onChange="TnSearchEvtSelect(this.value)">
				<option value="evt_code_search" selected>이벤트코드</option>
				<option value="evt_name_search">이벤트명</option>
				<option value="evt_partner_search">진행업체</option>
			</select>
			<input type="text" name="evt_code_search" id="evt_code_search" value="<%= evt_code_search%>" size="20" class="text">
			<input type="text" name="evt_name_search" id="evt_name_search" value="<%= evt_name_search%>" size="20" class="text" style="display:none">
			<input type="text" name="evt_partner_search" id="evt_partner_search" value="<%= evt_partner_search%>" size="20" class="text" style="display:none">
		</td>
	</tr>
	</form>
</table>
<!-- 검색 끝 -->
<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<tr>
		<td align="left"></td>
		<td align="right">
			<select class="select" onchange="document.frm.sortMtd.value=this.value;document.frm.submit();">
				<option value="">등록순</option>
				<option value="ws" <%=chkIIF(sortMtd="ws","selected","")%>>웹 정렬순</option>
				<option value="ms" <%=chkIIF(sortMtd="ms","selected","")%>>모바일 정렬순</option>
			</select>
		</td>
	</tr>
</table>
<!-- 액션 끝 -->

<!-- 리스트 시작 -->
<form action="" name="frmBuyPrc" method="POST" style="margin:0px;">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="evt_code_search" value="<%= evt_code_search %>">
<input type="hidden" name="evt_name_search" value="<%= evt_name_search %>">
<input type="hidden" name="evt_partner_search" value="<%= evt_partner_search %>">
<input type="hidden" name="evt_type_searchbox" value="<%= evt_type_search %>">
<input type="hidden" name="isusing_searchbox" value="<%= isusing_search %>">
<input type="hidden" name="evt_code_countbox" value="<%= evt_code_count %>">
<input type="hidden" name="evt_mobile_yn" value="<%= evt_mobile_yn %>">
<input type="hidden" name="page" value="<%= page %>">
<input type="hidden" name="sortMtd" value="<%=sortMtd%>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<% if oip.FresultCount>0 then %>
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="9">
			검색결과 : <b><%= oip.FTotalCount %></b>
			&nbsp;
			페이지 : <b><%= page %>/ <%= oip.FTotalPage %></b>
		</td>
		<td colspan="7" align="right">
			범주:
			<span style="padding:0 3px;border:1px #ccc solid;background-color:#fff;">현재 진행중</span>&nbsp;
			<span style="padding:0 3px;border:1px #ccc solid;background-color:#cfc;">당첨자O/오픈</span>&nbsp;
			<span style="padding:0 3px;border:1px #ccc solid;background-color:#fea;">당첨자X/종료</span>&nbsp;
			<span style="padding:0 3px;border:1px #ccc solid;background-color:#fcc;">최종종료</span>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td align="center" >이벤트 코드</td>
		<td align="center">이미지</td>
		<td align="center">이벤트 타입</td>
		<td align="center">이벤트명</td>
		<td align="center">진행업체</td>
		<td align="center">시작일</td>
		<td align="center">종료일</td>
		<td align="center">발표일</td>
		<td align="center">사용</td>
		<td align="center">코맨트수</td>
		<td align="center">마케팅담당</td>
		<td align="center">WD담당</td>
    </tr>
	<% for i=0 to oip.FresultCount-1 %>
    <tr align="center" bgcolor="#FFFFFF">
			<td align="center">
				<% if gubun="MC" then %>
				<a href="" onclick="jsSetEvtContMobile('<%= oip.FItemList(i).fevt_code %>');return false;">
				<% else %>
				<a href="" onclick="jsSetEvtCont('<%= oip.FItemList(i).fevt_code %>');return false;">
				<% end if %>
				<%= oip.FItemList(i).fevt_code %><br/>[선택]</a>
			</td>
			<td align="center">
				<image src="<%=webImgUrl%>/culturestation/2009/list/<%= oip.FItemList(i).fimage_list %>" width="40" height="40" border=0>
			</td>
			<td align="center">
			<% if oip.FItemList(i).fevt_type = "0" then
					response.write "느껴봐"
				elseif oip.FItemList(i).fevt_type = "1" then
					response.write "읽어봐"
				else
					response.write "들어봐"
				end if%></td>
			<td align="center">
				<%= oip.FItemList(i).fevt_name %>
			</td>
			<td align="center">
				<%= oip.FItemList(i).fevt_partner %>
			</td>
			<td align="center"><%= left(oip.FItemList(i).fstartdate,10) %></td>
			<td align="center"><%= left(oip.FItemList(i).fenddate,10) %></td>
			<td align="center"><%= left(oip.FItemList(i).feventdate,10) %></td>
			<td align="center"><%= "<span title='이벤트 사용여부'>" & oip.FItemList(i).fisusing & "</span>" %></td>
			<td align="center">
				<% if oip.FItemList(i).fevt_code_count = 0 then %>
				0
				<% else %>
					<a href="javascript:comment_list(<%= oip.FItemList(i).fevt_code %>);" onfocus="this.blur()">
					<%= oip.FItemList(i).fevt_code_count %><br>[보기]</a>
				<% end if %>
			</td>
			<td align="center"><%= oip.FItemList(i).femName %></td>
			<td align="center"><%= oip.FItemList(i).fedName %></td>
    </tr>
	<% next %>
	<% else %>
		<tr bgcolor="#FFFFFF">
			<td colspan="16" align="center" class="page_link">[검색결과가 없습니다.]</td>
		</tr>
	<% end if %>
    <tr height="25" bgcolor="FFFFFF">
		<td colspan="16" align="center">
	       	<% if oip.HasPreScroll then %>
				<span class="list_link"><a href="javascript:goPage(<%= oip.StartScrollPage-1 %>)">[pre]</a></span>
			<% else %>
			[pre]
			<% end if %>
			<% for i = 0 + oip.StartScrollPage to oip.StartScrollPage + oip.FScrollCount - 1 %>
				<% if (i > oip.FTotalpage) then Exit for %>
				<% if CStr(i) = CStr(oip.FCurrPage) then %>
				<span class="page_link"><font color="red"><b><%= i %></b></font></span>
				<% else %>
				<a href="javascript:goPage(<%= i %>)" class="list_link"><font color="#000000"><%= i %></font></a>
				<% end if %>
			<% next %>
			<% if oip.HasNextScroll then %>
				<span class="list_link"><a href="javascript:goPage(<%= i %>)">[next]</a></span>
			<% else %>
			[next]
			<% end if %>
		</td>
	</tr>
</table>
</form>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
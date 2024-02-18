<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 업체 입점문의
' History : 서동석 생성
'			2022.09.13 한용민 수정(엑셀다운로드,검색조건 추가)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/classes/board/companyrequestcls.asp" -->
<!-- #include virtual="/lib/classes/displaycate/displaycateCls.asp"-->
<%
dim i, j,ix, page,gubun, onlymifinish, license_no, research, searchkey,catevalue, dispCate,maxDepth
dim ipjumYN , catemid ,catelarge, sellgubun, workid, iid, reqcomment, startdate, enddate
	page 			= requestCheckvar(getNumeric(request("pg")),10)
	gubun 			= requestCheckvar(request("gubun"),2)
	onlymifinish 	= requestCheckvar(request("onlymifinish"),3)
	research 		= requestCheckvar(request("research"),3)
	searchkey 		= requestCheckvar(request("searchkey"),32)
	catevalue		= requestCheckvar(request("catevalue"),3)
	ipjumYN			= requestCheckvar(request("ipjumYN"),1)
	catemid 		= requestCheckvar(request("catemidbox"),3)
	catelarge 		= requestCheckvar(request("catelargebox"),3)
	dispCate		= requestCheckVar(Request("disp"),16) 
	maxDepth		= 2
	sellgubun			= requestCheckvar(request("sellgubun"),1)
	workid			= requestCheckvar(request("workid"),34)
	iid             = requestCheckVar(Request("iid"),9) 
	license_no		= requestCheckvar(request("license_no"),50)
	reqcomment		= requestCheckvar(request("reqcomment"),50)
	startdate = NullFillWith(requestCheckVar(request("startdate"),10),DateAdd("m",-1,date()))
	enddate = NullFillWith(requestCheckVar(request("enddate"),10),date())

'// 기본값으로 입점의뢰서
if gubun="" then gubun="01"
if research="" and onlymifinish="" then onlymifinish="on"		
if (page = "") then page = "1"
If gubun = "01" Then 
	'sellgubun = ""
	workid = ""
End If

dim companyrequest
set companyrequest = New CCompanyRequest
	companyrequest.PageSize = 20
	companyrequest.CurrPage = CInt(page)
	companyrequest.ScrollCount = 10
	companyrequest.FReqcd=gubun
	companyrequest.FOnlyNotFinish = onlymifinish
	companyrequest.FRectSearchKey = searchkey
	companyrequest.FRectCatevalue = catevalue
	companyrequest.FipjumYN = ipjumYN
	companyrequest.FRectDispCate = dispCate
	companyrequest.FRectSellgubun = sellgubun
	companyrequest.FRectWorkid = workid
	companyrequest.FRectID=iid
	companyrequest.FRectlicense_no=license_no
	companyrequest.FRectReqcomment=reqcomment
	companyrequest.FRectstartdate=startdate
	companyrequest.FRectenddate=DateAdd("d",+1,enddate)
	companyrequest.list

%>
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript">

function checkComp(comp){
	if(comp.value == '02'){
		document.location.href='/admin/board/upche/req_list.asp?menupos=<%=menupos%>&gubun=02&disp=&catevalue=';
	}else if(comp.value == '01'){
		document.location.href='/admin/board/upche/req_list.asp?menupos=<%=menupos%>';
	}
}

//프린트
function printpage(id){
	
	var printpage;
	printpage = window.open("/admin/board/upche/req_print.asp?id=" +id, "printpage","width=1024,height=768,scrollbars=yes,resizable=yes");
	printpage.focus();

}

function delitem(id){
	
	if (confirm("삭제하시겠습니까?.") ==true)
		frmdel.mode.value="del";
		frmdel.id.value=id;
		frmdel.submit();
}
function MovePage(page){
	frm.target="";
	frm.pg.value=page;
	frm.research.value="<%=research %>";
	//frm.gubun.value="<%=gubun%>";
	frm.onlymifinish.value="<%=onlymifinish%>";
	frm.catevalue.value="<%=catevalue%>";
	frm.ipjumYNvalue="<%=ipjumYN%>";
	frm.searchkey.value="<%=searchkey%>";
	frm.action="/admin/board/upche/req_list.asp";
	frm.submit();
}

function ViewPage(id){
	<% if gubun="02" then %>
		var winView = window.open("/admin/board/upche/req_view2.asp?id="+id,"popReq","width=1400,height=768,scrollbars=yes,resizable=yes");
	<% else %>
		var winView = window.open("/admin/board/upche/req_view.asp?id="+id,"popReq","width=1400,height=768,scrollbars=yes,resizable=yes");
	<% end if %>
	winView.focus();
/*
	var winView = window.open("about:blank;","popReq","width=1024,height=768,scrollbars=yes,resizable=yes");
	frm.id.value=id;
	frm.pg.value=<%=page%>;
	frm.research.value="<%=research %>";
	//frm.gubun.value="<%=gubun%>";
	frm.onlymifinish.value="<%=onlymifinish%>";
	frm.catevalue.value="<%=catevalue%>";
	frm.ipjumYNvalue="<%=ipjumYN%>";
	frm.searchkey.value="<%=searchkey%>";
	frm.target = "popReq";
	<% if gubun="02" then %>
		frm.action="/admin/board/upche/req_view2.asp";
	<% else %>
		frm.action="/admin/board/upche/req_view.asp";
	<% end if %>
	frm.action="";
	frm.target="";
	frm.submit();
*/
}

function DownPage(id,sFN){
	  var winFD = window.open("<%=uploadImgUrl%>/linkweb/company/downcorequest.asp?idx="+id+"&sFN="+sFN,"popFD","");
    winFD.focus();
} 

function changecontent() {
	frm.pg.value="1";
	frm.action="";
	frm.target="";
	frm.submit();
}

function downloadexcel(){
	document.frm.target = "view";
	document.frm.action = "/admin/board/upche/req_list_excel.asp";
	document.frm.submit();
	document.frm.target = "";
	document.frm.action = "";
}

</script> 

<!-- 검색 시작 -->
<form name="frm" method="get" action="" style="margin:0px;">
<input type="hidden" name="id" value="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="research" value="on">
<input type="hidden" name="pg" value="">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
<% If gubun <> "02" Then %>
	채널 : 
	<select name="sellgubun" class="select">
		<option value="" <%= Chkiif(sellgubun="", "selected", "") %> >전체</option>
		<option value="Y" <%= Chkiif(sellgubun="Y", "selected", "") %> >온라인/오프라인</option>
		<option value="N" <%= Chkiif(sellgubun="N", "selected", "") %> >온라인</option>
		<option value="F" <%= Chkiif(sellgubun="F", "selected", "") %> >오프라인</option>
	</select>
	/ 관리카테고리 
	<% call DrawSelectBoxCategoryLarge("catevalue",catevalue) %>
	/ 전시카테고리 <!-- #include virtual="/common/module/dispCateSelectBoxDepth.asp"-->&nbsp;
<% Else %>
	<input type="hidden" name="catevalue" value="">
	제휴종류 : 
	<select name="sellgubun" class="select">
		<option value="">전체</option>
		<option value="1" <%= Chkiif(sellgubun="1", "selected", "") %> >공급제휴</option>
		<option value="2" <%= Chkiif(sellgubun="2", "selected", "") %> >컨텐츠제휴</option>
		<option value="3" <%= Chkiif(sellgubun="3", "selected", "") %> >공동마케팅 및 프로모션 제휴</option>
		<option value="4" <%= Chkiif(sellgubun="4", "selected", "") %> >문화이벤트 제휴</option>
		<option value="5" <%= Chkiif(sellgubun="5", "selected", "") %> >기술 및 솔루션 관련 제휴</option>
		<option value="6" <%= Chkiif(sellgubun="6", "selected", "") %> >광고문의</option>
	</select>&nbsp;&nbsp;
	담당자 : 
	<% DrawWorkIdCombo "workid", workid %>
<% End If %>
	<select name="ipjumYN" class="a">
		<option value="">완료구분</option>
		<option value="Y" <% if ipjumYN="Y" then response.write "selected" %>>입점완료</option>
		<option value="N" <% if ipjumYN="N" then response.write "selected" %>>미완료</option>
	</select>
	</td>	
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="changecontent();">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
	<label><input type="radio" name="gubun" value="01" onclick="checkComp(this)" <% if gubun="01" then response.write "checked" %> >입점의뢰서</label>
	<label><input type="radio" name="gubun" value="02" onclick="checkComp(this)" <% if gubun="02" then response.write "checked" %> >사업제휴서</label>
	<!--<input type="radio" name="gubun" value="03" <% if gubun="03" then response.write "checked" %> >특정상품의뢰-->
	<!--<input type="radio" name="gubun" value="04" <% if gubun="04" then response.write "checked" %> >추천상품의뢰-->
	&nbsp;&nbsp;<label><input type="checkbox" name="onlymifinish" <% if onlymifinish="on" then response.write "checked" %> >처리안된목록</label>
	<br>
	<label>업체명 : <input type="text" name="searchkey" value="<%= searchkey %>" /></label>
	&nbsp;
	<label>글번호 : <input type="text" name="iid" value="<%= iid %>" size="6" /></label>
	&nbsp;
	<label>사업자등록번호 : <input type="text" name="license_no" value="<%= license_no %>" size="10" maxlength="50" /></label>
	<% if gubun="01" then %>
	&nbsp;
	<label>상품명(브랜드명) : <input type="text" name="reqcomment" value="<%= reqcomment %>" size="10" maxlength="50" /></label>
	<% end if %>
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		신청일 :
		<input type="text" name="startdate" id="startdate" value="<%= startdate %>" style="text-align:center;height:35px;" size="10" maxlength="10" readonly>
		<strong>&nbsp;~&nbsp;</strong>
		<input type="text" name="enddate" id="enddate" value="<%= enddate %>" style="text-align:center;height:35px;" size="10" maxlength="10" readonly>
		<script type="text/javascript">
			var CAL_Start = new Calendar({
				inputField : "startdate", trigger    : "startdate",
				onSelect: function() {
					var date = Calendar.intToDate(this.selection.get());
					CAL_End.args.min = date;
					CAL_End.redraw();
					this.hide();
				}, bottomBar: true, dateFormat: "%Y-%m-%d"
			});
			var CAL_End = new Calendar({
				inputField : "enddate", trigger    : "enddate",
				onSelect: function() {
					var date = Calendar.intToDate(this.selection.get());
					CAL_Start.args.max = date;
					CAL_Start.redraw();
					this.hide();
				}, bottomBar: true, dateFormat: "%Y-%m-%d"
			});
		</script>
	</td>
</tr>
</table>
</form>
<!-- 검색 끝 -->
<Br>
<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left"></td>
	<td align="right">
		<input type="button" onclick="downloadexcel();" value="엑셀다운로드" class="button">
	</td>
</tr>
</table>
<!-- 액션 끝 -->

<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		검색결과 : <b><%= companyrequest.TotalCount %></b>
		&nbsp;
		페이지 : <b><%= page %>/ <%= companyrequest.TotalPage %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
<td align="center">번호</td>
<td align="center">신청일</td>
<td align="center">제목</td>
<td align="center">채널</td>
<td align="center">처리일</td>
<td align="center">입점여부</td>
<td align="center">카테고리구분</td>
<td align="center">회사URL</td>
<td align="center">답변여부</td>
	<td align="center">비고</td>
</tr>
<% if companyrequest.resultCount>0 then %>
	<% for i = 0 to (companyrequest.ResultCount - 1) %>
	<tr align="center" bgcolor="#FFFFFF">
		<td><%=companyrequest.results(i).id%></td>
		<td align="center" nowrap><%= FormatDate(companyrequest.results(i).regdate, "0000-00-00") %></td>
		<td>[<%= companyrequest.code2name(companyrequest.results(i).reqcd) %>] <%= companyrequest.results(i).companyname %></td>
		<td align="center">
			<% if companyrequest.results(i).sellgubun="Y" then %>온라인/오프라인<%
			elseif companyrequest.results(i).sellgubun="N" then %>온라인<%
			elseif companyrequest.results(i).sellgubun="F" then %>오프라인<%
			else %><%=companyrequest.results(i).sellgubun%><%
			end if %>
		</td>
		<td align="center" nowrap>
			<% if (IsNull(companyrequest.results(i).finishdate) = true) then %>
			<font color="red">미완료</font>
			<% else %>
			<%= FormatDate(companyrequest.results(i).finishdate, "0000-00-00") %>
			<% end if %>
		</td>
		<td align="center">
			<%if companyrequest.results(i).ipjumYN="Y" then response.write "입점완료" %>
			<%if companyrequest.results(i).ipjumYN="N" then response.write "N" %>
			</td>
		<td align="center" nowrap>
			<div><%IF not isNull(companyrequest.results(i).dispcate) THEN%><%=companyrequest.results(i).dispcatename1%> > <%=companyrequest.results(i).dispcatename2%><%END IF%></div>
			<div style="color:gray"><%=companyrequest.results(i).cd1name%></div>  
			</td>
		<td align="center">
			<a href="<%IF left(companyrequest.results(i).companyurl,4)<>"http" then%>http://<%END IF%><%= companyrequest.results(i).companyurl%>" target="_blank"><%= companyrequest.results(i).companyurl%></a>
		</td> 
		<td align="center">
			<% if companyrequest.commentcheck(companyrequest.results(i).replycomment)="Y" then %>
			Y
			<% else %>
			<font color="red">N</font>
			<% end if %>
		</td>		
		<td align="center" nowrap>
			<input type="button" value="보기" class="button" onclick="javascript:ViewPage(<%= companyrequest.results(i).id %>);">
			<% if gubun="01" then %><input type="button" value="프린트" class="button" onclick="javascript:printpage(<%= companyrequest.results(i).id %>);"><% end if %>
			<%if companyrequest.results(i).attachfile <> "" then%><input type="button" value="첨부파일다운" class="button" onclick="javascript:DownPage(<%= companyrequest.results(i).id %>,'<%=companyrequest.results(i).attachfile%>');"><%end if%>
		</td>
	</tr>
	<% next %>
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15" align="center">
		<% if companyrequest.HasPreScroll then %>
			<a href="javascript:MovePage(<%= companyrequest.StartScrollPage-1 %>);">[prev]</a>
		<% else %>
			[prev]
		<% end if %>

		<% for ix=0 + companyrequest.StartScrollPage to companyrequest.ScrollCount + companyrequest.StartScrollPage - 1 %>
			<% if ix>companyrequest.Totalpage then Exit for %>
				<% if CStr(page)=CStr(ix) then %>
					<font color="red">[<%= ix%>]</font>
				<% else %>
					<a href="javascript:MovePage(<%=ix%>);">[<%= ix %>]</a>
				<% end if %>
		<% next %>

		<% if companyrequest.HasNextScroll then %>
			<a href="javascript:MovePage(<%=ix%>);">[next]</a>
		<% else %>
			[next]
		<% end if %>
		</td>
	</tr>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="10" align="center" class="page_link">[검색결과가 없습니다.]</td>
	</tr>
<% end if %>
</table>

<form name="frmdel" method="get" action="/admin/board/upche/cscenter_req_board_act.asp" style="margin:0px;">
<input type="hidden" name="mode" value="">
<input type="hidden" name="id" value="">
<input type="hidden" name="page" value="<%=page%>">
</form>

<% IF application("Svr_Info")="Dev" THEN %>
	<iframe id="view" name="view" src="" width="100%" height=300 frameborder="0" scrolling="no"></iframe>
<% else %>
	<iframe id="view" name="view" src="" width=0 height=0 frameborder="0" scrolling="no"></iframe>
<% end if %>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
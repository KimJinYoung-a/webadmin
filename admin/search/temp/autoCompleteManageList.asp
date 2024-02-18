<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"

If not (Request.ServerVariables("REMOTE_ADDR") = "61.252.133.75" or Request.ServerVariables("REMOTE_ADDR") = "61.252.133.105" or Request.ServerVariables("REMOTE_ADDR") = "61.252.133.106") Then
	Response.End
End If
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/lib/classes/admin/menucls.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/search/search_manageCls.asp"-->
<%
Dim i, cAuto, vPage, vSDate, vEDate, vAutoType, vSearchTxt, vUseYN
vPage = NullFillWith(requestCheckVar(Request("page"),10),1)
vSDate = requestCheckVar(Request("sdate"),10)
vEDate = requestCheckVar(Request("edate"),10)
vAutoType = requestCheckVar(Request("autotype"),2)
vUseYN = NullFillWith(requestCheckVar(Request("useyn"),1),"y")
vSearchTxt = requestCheckVar(Request("searchtxt"),50)

Set cAuto = New CSearchMng
cAuto.FCurrPage = vPage
cAuto.FPageSize = 15
cAuto.FRectSDate = vSDate
cAuto.FRectEDate = vEDate
cAuto.FRectAutoType = vAutoType
cAuto.FRectUseYN = vUseYN
cAuto.FRectSearchTxt = vSearchTxt
cAuto.fnAutoCompleteList

%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" lang="ko" xml:lang="ko">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<script language="JavaScript" src="/js/xl.js"></script>
<script language="JavaScript" src="/js/common.js"></script>
<script language="JavaScript" src="/js/report.js"></script>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script language="JavaScript" src="/js/calendar.js"></script>
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<script language="JavaScript" src="/cscenter/js/cscenter.js"></script>
<link rel="stylesheet" type="text/css" href="/css/adminDefault.css" />
<link rel="stylesheet" type="text/css" href="/css/adminCommon.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<!--[if IE]>
	<link rel="stylesheet" type="text/css" href="/css/adminPartnerIe.css" />
<![endif]-->
<script language='javascript'>
function searchFrm(p){
	frm1.page.value = p;
	frm1.submit();
}

function jsAutoReg(idx){
	var popautoreg;
	popautoreg = window.open('autoCompleteManage.asp?idx='+idx+'','popautoreg','width=800,height=815,scrollbars=yes,resizable=yes');
	popautoreg.focus();
}

function jsAllClick(){
	if($("#allclick").prop("checked")){
		$("input[name=idxarr]:checkbox").each(function() {
			$(this).prop("checked", true);
			jsThisClick($(this).val());
		});
	}else{
		$("input[name=idxarr]:checkbox").each(function() {
			$(this).prop("checked", false);
			jsThisCheck($(this).val());
		});
	}
}

function jsThisClick(i){
	$("#tr"+i+"").css('backgroundColor', '#D9FFFF');
	$("#idx"+i+"").prop("checked", true);
	jsThisCheck(i);
}

function jsThisCheck(i){
	if($("#idx"+i+"").is(":checked")){
		$("#tr"+i+"").css('backgroundColor', '#D9FFFF');
	}else{
		$("#tr"+i+"").css('backgroundColor', '#FFFFFF');
	}
}

function jsAutoProc(g){
	var msg;
	
	if(g == "update_arr"){
		msg = "변경";
	}else{
		msg = "삭제";
	}
	
	if($(":checkbox[name=idxarr]:checked").length == "0"){
		alert(""+msg+"할 자동완성을 선택하세요.");
		return;
	}
	
	if(confirm("선택된 자동완성 정보를 "+msg+"하시겠습니까?") == true) {
		$("#action").val(g);
		
		var tt;
		var ii = 1;
		$("input[name='idxarr']:checkbox:checked").each(function(){
			if(ii == 1){
				tt = $("#title"+$(this).val()).val();
			}else{
				tt = tt + "," + $("#title"+$(this).val()).val();
			}
			ii = ii + 1
		});
		$("#titlearr").val(tt);
		
		frm2.submit();
	}
}
</script>
</head>
<body>
<div id="calendarPopup" style="position: absolute; visibility: hidden; z-index: 2;"></div>
<div class="contSectFix scrl">
	<div class="contHead">
		<div class="locate"><h2><strong>자동완성 키워드 관리</strong></h2></div>
		<div class="helpBox">
			<a href="autoCompleteManageList.asp">자동완성 키워드 관리</a> l 
			<a href="quickLinkManageList.asp">퀵링크 관리</a> l 
			<a href="keywordQratingManageList.asp">키워드 큐레이터 관리</a>
		</div>
	</div>

	<!-- 상단 검색폼 시작 -->
	<form name="frm1" method="get" action="">
	<input type="hidden" name="page" value="<%=vPage%>">
	<!-- search -->
	<div class="searchWrap">
		<div class="search">
			<ul>
				<li>
					<label class="formTit">기간 :</label>
					<input type="text" class="formTxt" id="sdate" name="sdate" value="<%=vSDate%>" style="width:100px" placeholder="시작일" maxlength="10" readonly />
					<img src="/images/admin_calendar.png" id="sdate_trigger" alt="달력으로 검색" />
					<script language="javascript">
						var CAL_Start = new Calendar({
							inputField : "sdate", trigger    : "sdate_trigger",
							onSelect: function() {
								var date = Calendar.intToDate(this.selection.get());
								CAL_End.args.min = date;
								CAL_End.redraw();
								this.hide();
							}, bottomBar: true, dateFormat: "%Y-%m-%d"
						});
					</script>
					~
					<input type="text" class="formTxt" id="edate" name="edate" value="<%=vEDate%>" style="width:100px" placeholder="종료일" maxlength="10" readonly />
					<img src="/images/admin_calendar.png" id="edate_trigger" alt="달력으로 검색" />
					<script language="javascript">
						var CAL_End = new Calendar({
							inputField : "edate", trigger    : "edate_trigger",
							onSelect: function() {
								var date = Calendar.intToDate(this.selection.get());
								CAL_Start.args.max = date;
								CAL_Start.redraw();
								this.hide();
							}, bottomBar: true, dateFormat: "%Y-%m-%d"
						});
					</script>
				</li>
			</ul>
		</div>
		<dfn class="line"></dfn>
		<div class="search">
			<ul>
				<li>
					<p class="formTit">자동완성 속성 :</p>
					<select class="formSlt" id="autotype" title="옵션 선택" name="autotype">
						<option value="" <%=CHKIIF(vAutoType="","selected","")%>>전체</option>
						<%=fnAutoCompleteTypeSelect(vAutoType)%>
					</select>
				</li>
			</ul>
		</div>
		<dfn class="line"></dfn>
		<div class="search">
			<ul>
				<li>
					<p class="formTit">사용여부 :</p>
					<span class="rMar10"><input type="radio" id="useyny" name="useyn" value="y" <%=CHKIIF(vUseYN="y","checked","")%> /> <label for="useyny">사용함</label></span>
					<span class="rMar10"><input type="radio" id="useynn" name="useyn" value="n" <%=CHKIIF(vUseYN="n","checked","")%> /> <label for="useynn">사용안함</label></span>
				</li>
			</ul>
		</div>
		<dfn class="line"></dfn>
		<div class="search">
			<ul>
				<li>
					<label class="formTit" for="schWord">검색어 :</label>
					<input type="text" class="formTxt" id="searchtxt" name="searchtxt" value="<%=vSearchTxt%>" style="width:500px" placeholder="제목을 검색하세요." />
				</li>
			</ul>
		</div>
		<input type="button" class="schBtn" value="검색" onClick="searchFrm(1);" />
	</div>
	<!-- //search -->
	</form>

	<div class="cont">
		<div class="pad20">
			<div class="overHidden">
				<div class="ftLt">
					<input type="button" class="btn" value="자동완성 등록" onClick="jsAutoReg('');" />
					<% If cAuto.FResultCount > 0 Then %>
					<input type="button" class="btn bold" value="적용" onClick="jsAutoProc('update_arr');" />
					<input type="button" class="btn cGy1" value="삭제" onClick="jsAutoProc('delete_arr');" />
					<% End If %>
				</div>
			</div>

			<form name="frm2" action="autoCompleteProc.asp" method="post" target="iframeproc">
			<input type="hidden" id="action" name="action" value="">
			<input type="hidden" id="titlearr" name="titlearr" value="">
			<div>
				<div class="rt pad10">
					<span>검색결과 : <strong><%=FormatNumber(cAuto.FTotalCount,0)%></strong> <span class="lMar10">페이지 : <strong><%=cAuto.FtotalPage%> / <%=FormatNumber(vPage,0)%></strong></span>
				</div>
				<table class="tbType1 listTb">
					<thead>
					<tr>
						<th><div><% If cAuto.FResultCount > 0 Then %><input type="checkbox" id="allclick" onClick="jsAllClick();" /><% End If %></div></th>
						<th><div>No.</div></th>
						<th><div>자동완성 속성</div></th>
						<th><div>제목</div></th>
						<th><div>아이콘</div></th>
						<th><div>최종작성일(최종작업자)</div></th>
						<th><div>URL</div></th>
						<th><div></div></th>
					</tr>
					</thead>
					<tbody>
					<%
						If cAuto.FResultCount > 0 Then
							For i=0 To cAuto.FResultCount-1
					%>
							<tr id="tr<%=cAuto.FItemList(i).Fidx%>">
								<td><div><input type="checkbox" id="idx<%=cAuto.FItemList(i).Fidx%>" name="idxarr" value="<%=cAuto.FItemList(i).Fidx%>" onClick="jsThisCheck('<%=cAuto.FItemList(i).Fidx%>');" /></div></td>
								<td><%=cAuto.FItemList(i).Fidx%></td>
								<td>
									<%=fnAutoCompleteTypeName(cAuto.FItemList(i).Fautotype)%>
									<input type="hidden" name="autotypearr" value="<%=cAuto.FItemList(i).Fautotype%>">
								</td>
								<td><input type="text" class="formTxt" id="title<%=cAuto.FItemList(i).Fidx%>" name="title" value="<%=cAuto.FItemList(i).Ftitle%>" style="width:100%" onClick="jsThisClick('<%=cAuto.FItemList(i).Fidx%>');" placeholder="고객행복센터" maxlength="10" /></td>
								<td><%=fnAutoCompleteIconName(cAuto.FItemList(i).Ficon)%></td>
								<td><%=cAuto.FItemList(i).Flastusername%>(<%=cAuto.FItemList(i).Flastdate%>)</td>
								<td>
									<a href="http://www.10x10.co.kr<%=cAuto.FItemList(i).Furl_pc%>" class="cBl1 tLine" target="_blank">[PC바로가기]</a><br />
									<a href="http://m.10x10.co.kr<%=cAuto.FItemList(i).Furl_m%>" class="cBl1 tLine" target="_blank">[M바로가기]</a>
								</td>
								<td><input type="button" class="btn" value="수정" onClick="jsAutoReg('<%=cAuto.FItemList(i).Fidx%>');" /></td>
							</tr>
					<%
							Next
						End If
					%>
					</tfoot>
				</table>
				<div class="ct tPad20 cBk1">
					<% if cAuto.HasPreScroll then %>
					<a href="javascript:searchFrm('<%= cAuto.StartScrollPage-1 %>')">[pre]</a>
					<% else %>
		    			[pre]
		    		<% end if %>
		    		
		    		<% for i=0 + cAuto.StartScrollPage to cAuto.FScrollCount + cAuto.StartScrollPage - 1 %>
		    			<% if i>cAuto.FTotalpage then Exit for %>
		    			<% if CStr(vPage)=CStr(i) then %>
		    			<span class="cRd1">[<%= i %>]</span>
		    			<% else %>
		    			<a href="javascript:searchFrm('<%= i %>')">[<%= i %>]</a>
		    			<% end if %>
		    		<% next %>
					
					<% if cAuto.HasNextScroll then %>
		    			<a href="javascript:searchFrm('<%= i %>')">[next]</a>
		    		<% else %>
		    			[next]
		    		<% end if %>
				</div>
			</div>
			</form>
		</div>
	</div>
</div>
<iframe src="about:blank" name="iframeproc" width="0" height="0" frameborder="0"></iframe>
<% Set cAuto = Nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->
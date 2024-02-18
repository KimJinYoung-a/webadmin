<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description : [업체]게시판
' History : 2015.05.27 누군지모름 생성
'		  :	2016.01.13 한용민 수정
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/classes/board/upche_qnacls.asp" -->
<%
dim itemqanotinclude, sWorkerGubun, boardqna, page, i,ix, workergubuntype
dim SearchKey, SearchString, gubun, replyYn, usingYn, param, selDate, sDate, eDate
dim research, isRecent, sortBy
	SearchKey = Request("SearchKey")
	SearchString = Request("SearchString")
	gubun = Request("gubun")
	replyYn = Request("replyYn")
	usingYn = Request("usingYn")
	selDate		= requestCheckVar(Request("selDate"),1)
	sDate 		= requestCheckVar(Request("sDate"),10)
	eDate 		= requestCheckVar(Request("eDate"),10)
	sWorkerGubun = Request("workergubun")
	workergubuntype = requestCheckVar(Request("workergubuntype"),10)
	page = getNumeric(request("page"))
	research = requestCheckVar(Request("research"),2)
	isRecent = requestCheckVar(Request("isRecent"),1)
	sortBy = requestCheckVar(Request("sortBy"),2)

if page="" then page=1
if SearchKey="" then SearchKey="title"
if selDate="" then selDate="R"
if workergubuntype="" then
	sWorkerGubun=""
elseif workergubuntype="MY" then
	sWorkerGubun = session("ssBctId")
end if
if isRecent="" and research<>"on" then isRecent="Y"		'최근글표시(기본)
if sortBy="" then sortBy="rd"							'정렬기준값(기본:최신순)

param = "&research=on&SearchKey=" & SearchKey & "&SearchString=" & Server.URLencode(SearchString) & "&gubun=" & gubun & "&replyYn=" & replyYn & "&usingYn=" & usingYn & "&selDate=" & selDate & "&sDate=" & sDate & "&eDate=" & eDate & "&workergubuntype=" & workergubuntype & "&isRecent=" & isRecent & "&sortBy=" & sortBy

set boardqna = New CUpcheQnA
	boardqna.Fcurrpage = page
	boardqna.FPageSize = 30
	'boardqna.FCurrPage = 1
	boardqna.FRectGubun = gubun
	boardqna.FRectRelpy = replyYn
	boardqna.FRectUsing = usingYn
	boardqna.FRectSearchKey = SearchKey
	boardqna.FRectSearchString = SearchString
	boardqna.FWorkerGubun = sWorkerGubun
	boardqna.Frectworkergubuntype = workergubuntype
	boardqna.FRectSelDate = selDate
	boardqna.FRectSDate = sDate
	boardqna.FRectEDate = eDate
	boardqna.FRectIsRecenct = isRecent
	boardqna.FRectSortBy = sortBy
	boardqna.getqnalist

%>
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript">

function chworkergubuntype(chval){
	if (chval==''){
		$("#workergubun").hide();
	}else if (chval=='MY'){
		$("#workergubun").hide();
	}else if (chval=='SELECTID' || chval=='SELECTNAME'){
		<% if workergubuntype="MY" then %>
			$("#workergubun").val('');
		<% end if %>

		$("#workergubun").show();
	}
}

</script>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="research" value="on">
<input type="hidden" name="menupos" value="<%= menupos %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		키워드:
		<select class="select" name="SearchKey">
			<option value="title">제목</option>
			<option value="contents">내용</option>
			<option value="username">업체명</option>
            <option value="replyuser">답변자</option>
            <option value="userid">브랜드ID</option>
		</select>
		<input type="text" class="text" name="SearchString" size="12" value="<%=SearchString%>">
		&nbsp;
		기간:
		<select name="selDate" class="select">
			<option value="R" <%if Cstr(selDate) = "R" THEN %>selected<%END IF%>>작성일 기준</option>
			<option value="A" <%if Cstr(selDate) = "A" THEN %>selected<%END IF%>>답변일 기준</option>
		</select>
        <input id="sDate" name="sDate" value="<%=sDate%>" class="text" size="10" maxlength="10" /><img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="sDate_trigger" border="0" style="cursor:pointer" align="absmiddle" /> ~
        <input id="eDate" name="eDate" value="<%=eDate%>" class="text" size="10" maxlength="10" /><img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="eDate_trigger" border="0" style="cursor:pointer" align="absmiddle" />
		<script language="javascript">
			var CAL_Start = new Calendar({
				inputField : "sDate", trigger    : "sDate_trigger",
				onSelect: function() {
					var date = Calendar.intToDate(this.selection.get());
					CAL_End.args.min = date;
					CAL_End.redraw();
					this.hide();
				}, bottomBar: true, dateFormat: "%Y-%m-%d"
			});
			var CAL_End = new Calendar({
				inputField : "eDate", trigger    : "eDate_trigger",
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
		<input type="image" src="/admin/images/search2.gif" width="74" height="22" border="0" align="absmiddle">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		문의구분:
		<select class="select" name="gubun">
			<option value="">전체</option>
			<option value="01">배송문의</option>
			<option value="02">반품문의</option>
			<option value="03">교환문의</option>
			<option value="04">정산문의</option>
			<option value="05">입고문의</option>
			<option value="06">재고문의</option>
			<option value="07">상품등록문의</option>
			<option value="08">이벤트시행문의</option>
			<option value="20">기타문의</option>
		</select>
		&nbsp;
		답변여부:
		<select class="select" name="replyYn">
			<option value="">전체</option>
			<option value="Y">답변완료</option>
			<option value="N">미완료</option>
		</select>
		&nbsp;
		사용여부:
		<select class="select" name="usingYn">
			<option value="">사용</option>
			<option value="N">삭제</option>
		</select>
		<script language="javascript">
			document.frm.SearchKey.value="<%=SearchKey%>";
			document.frm.gubun.value="<%=gubun%>";
			document.frm.replyYn.value="<%=replyYn%>";
			document.frm.usingYn.value="<%=usingYn%>";
		</script>
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		담당자:
		<label><input type="radio" name="workergubuntype" value="" <% if workergubuntype="" then response.write " checked" %> onclick="chworkergubuntype('');">전체</label>
		<label><input type="radio" name="workergubuntype" value="MY" <% if workergubuntype="MY" then response.write " checked" %> onclick="chworkergubuntype('MY');">본인</label>
		<label><input type="radio" name="workergubuntype" value="SELECTID" <% if workergubuntype="SELECTID" then response.write " checked" %> onclick="chworkergubuntype('SELECTID');">ID검색</label>
		<label><input type="radio" name="workergubuntype" value="SELECTNAME" <% if workergubuntype="SELECTNAME" then response.write " checked" %> onclick="chworkergubuntype('SELECTNAME');">이름검색</label>
		<label><input type="text" class="text" name="workergubun" id="workergubun" size="8" value="<%= sworkergubun %>" style="display:none;">
		&nbsp;/&nbsp;
		정렬방법:
		<select class="select" name="sortBy">
			<option value="rd" <%=chkIIF(sortBy="rd","selected","")%>>최근등록순</option>
			<option value="ra" <%=chkIIF(sortBy="ra","selected","")%>>작성일순</option>
			<option value="ad" <%=chkIIF(sortBy="ad","selected","")%>>최근답변순</option>
		</select>
		&nbsp;/&nbsp;
		<label><input type="checkbox" name="isRecent" value="Y" <%=chkIIF(isRecent="Y","checked","")%>/>
		6개월 이내 검색</label>
	</td>
</tr>
</form>
</table>
<!-- 검색 끝 -->

<br>

<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="9">
		검색결과 : <b><%=FormatNumber(boardqna.FtotalCount,0)%></b>
		&nbsp;
		페이지 : <b><%= page %> / <%= boardqna.FtotalPage %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="50">idx</td>
	<td width="200">브랜드명(브랜드ID)</td>
    <td>제목</td>
    <td width="90">구분</td>
    <td width="70">담당자</td>
    <td width="70">작성일</td>
    <td width="70">답변여부</td>
    <td width="70">답변자</td>
   	<td width="60">답변일</td>
</tr>
<% for i = 0 to (boardqna.FResultCount - 1) %>
<tr height="25" align="center" <% if boardqna.FItemList(i).Fisusing="Y" then %>bgcolor="#FFFFFF"<% else %>bgcolor="#F2F2F2"<% end if %>>
	<td><%= boardqna.FItemList(i).Fidx %></td>
	<td><a href="/admin/board/upche_qna_board_reply.asp?idx=<%= boardqna.FItemList(i).Fidx & "&page=" & page & Param %>"><%= boardqna.FItemList(i).Fusername %>(<%= boardqna.FItemList(i).Fuserid %>)</a></td>
	<td align="left"><a href="/admin/board/upche_qna_board_reply.asp?idx=<%= boardqna.FItemList(i).Fidx & "&page=" & page & Param %>"><%= CHKIIF(boardqna.FItemList(i).Ftitle="", "(제목없음)", ReplaceBracket(db2html(boardqna.FItemList(i).Ftitle))) %></a></td>
	<td><a href="/admin/board/upche_qna_board_reply.asp?idx=<%= boardqna.FItemList(i).Fidx & "&page=" & page & Param %>"><%= boardqna.FItemList(i).GubunName %></a></td>
	<td><%= boardqna.FItemList(i).Fworker %></td>
	<td><%= FormatDate(boardqna.FItemList(i).Fregdate, "0000-00-00") %></td>
	<td>
		<% if not isnull(boardqna.FItemList(i).Freplyuser) then %>
			답변완료
		<% else %>
			&nbsp;
		<% end if %>
	</td>
	<td><%= boardqna.FItemList(i).Freplyuser %></td>
	<td><%= left(boardqna.FItemList(i).Freplydate,10) %></td>
</tr>
<% next %>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="9" align="center">
		<% if boardqna.HasPreScroll then %>
			<a href="?page=<%= boardqna.StartScrollPage-1 & param %>">[pre]</a>
		<% else %>
			[pre]
		<% end if %>
		<% for ix=0 + boardqna.StartScrollPage to boardqna.FScrollCount + boardqna.StartScrollPage - 1 %>
			<% if ix>boardqna.FTotalpage then Exit for %>
			<% if CStr(page)=CStr(ix) then %>
			<font color="red">[<%= ix %>]</font>
			<% else %>
			<a href="?page=<%= ix & param %>">[<%= ix %>]</a>
			<% end if %>
		<% next %>

		<% if boardqna.HasNextScroll then %>
			<a href="?page=<%= ix & param %>">[next]</a>
		<% else %>
			[next]
		<% end if %>
	</td>
</tr>
</table>

<script type="text/javascript">

chworkergubuntype('<%= workergubuntype %>')

</script>

<%
set boardqna = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->

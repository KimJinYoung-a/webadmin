<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/search/adminsearchCls.asp" -->
<%
 
dim ret, i, j, k
dim colNum, rowNum
dim resultKeyword, resultTag
dim serverAddr, pageSize

pageSize = requestCheckvar(request("pageSize"),10)
if (pageSize="") then pageSize=20
    
'dim SVRARR : SVRARR = array(G_ORGSCH_ADDR,G_1STSCH_ADDR,G_2NDSCH_ADDR,G_3RDSCH_ADDR,G_4THSCH_ADDR) 
'dim SVRInfo : SVRInfo = array("인덱싱","WWW","WWW(2)","APP","MOBILE")
dim SVRARR : SVRARR = array(G_ORGSCH_ADDR,G_2NDSCH_ADDR,G_1STSCH_ADDR,G_4THSCH_ADDR,G_3RDSCH_ADDR)
dim SVRInfo : SVRInfo = array("인덱싱","WWW(카테고리 내)","WWW","MOBILE","APP")

dim SVRCNT : SVRCNT = UBOUND(SVRARR)
Redim serverAddrArr(SVRCNT), resultKeywordArr(SVRCNT), resultTagArr(SVRCNT)

for i=0 to SVRCNT
    serverAddrArr(i) = SVRARR(i)   
next


dim osearch
set osearch= New SearchItemCls

'' =============================================================================
'// k 번 서버

for k = LBound(serverAddrArr) to UBOUND(serverAddrArr)
	serverAddr = serverAddrArr(k)

	''실시간 인기검색어
	osearch.FPageSize = pageSize
	resultKeyword = ""
	resultTag = ""
	ret = osearch.getRealtimePopularKeyWords(resultKeyword, resultTag, serverAddr, 1, 0)   ''검색어,태그,서버,실시간여부(메모리),도메인
	resultKeywordArr(k) = resultKeyword
	resultTagArr(k) = resultTag
next


'' =============================================================================

set osearch = Nothing

%>
<script language='javascript'>

/*
function NextPage(page){
    document.frm.page.value = page;
    document.frm.submit();
}

function jsPopRelatedKeywordAdd() {
    var popwin = window.open('popRelatedKeywordAdd.asp','jsPopRelatedKeywordAdd','width=330,height=220,scrollbars=yes,resizable=yes');
    popwin.focus();
}

function jsDelRelatedKeyword(idx) {
	var ret = confirm("삭제하시겠습니까?");
	if(ret){
		var frm = document.frmAct;
		frm.mode.value = "del";
		frm.idx.value = idx;
		frm.submit();
	}
}
*/

</script>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="research" value="on">
	<input type="hidden" name="page" value="1">
	<input type="hidden" name="menupos" value="<%= request("menupos") %>">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
		<td align="left" height="30" >
		    갯수:
		    <select name="pageSize">
		        <option value="20" <%=CHKIIF(pageSize="20","selected","")%> >20
		        <option value="30" <%=CHKIIF(pageSize="30","selected","")%> >30
		        <option value="50" <%=CHKIIF(pageSize="50","selected","")%> >50
		        <option value="100" <%=CHKIIF(pageSize="100","selected","")%> >100
		    </select>
		    &nbsp;/&nbsp;
			현재시간 : <%= Now() %>
			&nbsp;/&nbsp;
			순위 변경 - 5분단위
		</td>
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button" value="검색" onClick="javascript:document.frm.submit();">
		</td>
	</tr>
	</form>
</table>
<!-- 검색 끝 -->

<p>

<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
 	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<% for k = LBound(serverAddrArr) to UBound(serverAddrArr) %>
		<td>서버 <%= Format00(2, (k+1)) %>-<%= svrInfo(k) %><br>(<%= RIGHT(serverAddrArr(k),3) %>)</td>
		<% next %>
	</tr>

	<tr align="center" bgcolor="#FFFFFF">
		<% for k = LBound(serverAddrArr) to UBound(serverAddrArr) %>
		<td align="center" height="30">

			<% if isArray(resultKeywordArr(k)) then %>
			<table width="200" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
				<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
					<td width="100">검색어</td>
					<td>순위</td>
				</tr>
				<% for i = LBound(resultKeywordArr(k)) to UBound(resultKeywordArr(k)) %>
				<tr align="center" bgcolor="#FFFFFF">
					<td align="left"><%= resultKeywordArr(k)(i) %></td>
					<td align="left"><%= resultTagArr(k)(i) %></td>
				</tr>
				<% next %>
			</table>
			<% else %>
			ERR
			<% end if %>

		</td>
		<% next %>
	</tr>
</table>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->

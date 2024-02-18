<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/db/db2open.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->

<%

dim eventidx
eventidx= request("eventidx")

if eventidx="" then
response.write "<h1>잘못된 접근입니다</h1>"
dbget.close()	:	response.End
end if

		dim pagesize,page,searchvalue ,ScrollCount

				pagesize=35
				ScrollCount=10
				searchvalue= request("searchvalue")

				page= request("page")
				if page="" then page=1

		dim FTotalCount,FTotalPage,FResultCount,ix
		dim searchSQL,sql,strHtml

		if searchvalue<>"" then
			searchSQL = " and comment like '%" & searchvalue & "%' "
		end if



		SQL =	" select count(com_idx) as cnt from [db_cts].[dbo].[tbl_2007_diary_event_comment]" &_
					" where eventidx=" & eventidx & searchSQL

		db2_rsget.open sql, db2_dbget, 1

				if not db2_rsget.eof then
					FTotalCount = db2_rsget("cnt")
				end if

				db2_rsget.close

				FTotalPage =  CInt(FTotalCount\PageSize)

				if  (FTotalCount\PageSize)<>(FTotalCount/PageSize) then
					FTotalPage = FTotalPage +1
				end if


		SQL = " select top " & PageSize &_
					" com_idx,userid,comment " &_
					" from [db_cts].[dbo].[tbl_2007_diary_event_comment] " &_
					" where isusing='Y' " &_
					" and eventidx=" & eventidx &_
					searchSQL &_
					" and com_idx <=  " &_
					"							(select ISNULL(min(t.com_idx) ,0) " &_
					"							from (select top " & (PageSize*(Page-1))+1 & " com_idx  " &_
					"										from [db_cts].[dbo].[tbl_2007_diary_event_comment] " &_
					"										where eventidx =" & eventidx & " and isusing='Y' " & searchSQL & " order by com_idx desc) as t ) " &_
					" order by com_idx desc "


		db2_rsget.open SQL,db2_dbget,1

				if not db2_rsget.eof then
					FResultCount = db2_rsget.recordCount
				end if


		if not db2_rsget.eof then

				do until db2_rsget.eof
					strHtml = strHtml + "<tr bgcolor='#FFFFFF'>"
					strHtml = strHtml + "	<td width='80' align='center'>" & db2_rsget("com_idx") & "</td> "
					strHtml = strHtml + "	<td width='100' align='center'><span onclick=" & chr(34) & "DoWinner('" & db2_rsget("userid") & "')" & chr(34) & " style='cursor:pointer'>" & db2_rsget("userid") & "</span></td> "
					strHtml = strHtml + "	<td align='center'>" & db2_rsget("comment") & "</td> "
					'strHtml = strHtml + "	<td align='center'><span onclick=" & chr(34) & "TnGoWinnerList('" & db2_rsget("eventidx") & "')" & chr(34) & " style='cursor:pointer'>보기</span></td> "
					strHtml = strHtml + "</tr>"

				db2_rsget.movenext
				loop

		end if

		db2_rsget.close

	public Function HasPreScroll()
		HasPreScroll = StartScrollPage > 1
	end Function

	public Function HasNextScroll()
		HasNextScroll = FTotalPage > StartScrollPage + ScrollCount -1
	end Function

	public Function StartScrollPage()
		StartScrollPage = ((page-1)\ScrollCount)*ScrollCount +1
	end Function
%>

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<script language="JavaScript" src="/js/xl.js"></script>
<script language="JavaScript" src="/js/common.js"></script>
<script language="JavaScript" src="/js/report.js"></script>
<link rel="stylesheet" href="/bct.css" type="text/css">
</head>
<body leftmargin="0" topmargin="0">
<table width="699" border="0" cellpadding="0" cellspacing="1" bgcolor="#CCCCCC" class="a">
	<tr>
		<td colspan="3">
			<table width="690" border="0" cellpadding="0" cellspacing="0" bgcolor="#CCCCCC" class="a">
				<form name="searchFrm" method="post" action="">
				<tr>
					<td><a href="/admin/sitemaster/diary_collection_2007/diary_quiz_event_main.asp" target="_parent"><font color="blue">메인으로</font></a></td>
					<td align="right">
						정답자만 보기<input type="checkbox" name="checkRight" <% if searchvalue<>"" then response.write "checked" %> onclick="CheckSearch();"/>
						정답입력:<input type="text" name="searchInput" value="<%= searchvalue %>" readOnly=true />
						<input type="button" value="검색" onclick="FnSearch();" onKeyPress="if (event.keyCode == 13) FnSearch();"/>
					</td>
				<tr>
				</form>
			</table>
		</td>
	</tr>
	<tr bgcolor="#EDEDED">
		<td width="80" align="center">번호</td>
		<td width="100" align="center">아이디</td>
		<td width="600" align="center">코멘트</td>
	</tr>
	<%= strHtml %>

	<tr>
		<td colspan="3" align="center">
			<table width="300" border="0" cellpadding="0" cellspacing="0" class="a">
				<tr>
				<% if HasPreScroll then %>
					<td  style="padding:2 0 0 0"><a href="javascript:FnPageMove('<%= CStr(StartScrollPage - 1) %>');"> [prev] </a></td>
				<% else %>
					<td  style="padding:2 0 0 0"> [prev] </td>
				<% end if %>

				<% for ix = StartScrollPage to (StartScrollPage + ScrollCount - 1) %>
					<% if (ix > FTotalPage) then Exit For %>
					<% if CStr(ix) = CStr(page) then %>
					<td> |<font color="#DD0000"><strong><%= ix %></strong></font></td>
					<% else %>
					<td> |<a href="javascript:FnPageMove('<%= ix %>');" class="link_e3" onfocus="this.blur();"><%= ix %> </a></td>
					<% end if %>
				<% next %>

				<% if HasNextScroll then %>
					<td  style="padding:2 0 0 0"><a href="javascript:FnPageMove('<%= CStr(StartScrollPage + ScrollCount) %>');"> [next]</a></td>
				<% else %>
					<td  style="padding:2 0 0 0"> [next] </td>
				<% end if %>
				</tr>
			</table>
		</td>
	</tr>
</table>
</body>
</html>

<form name="pageFrm" method="post" action="?">
<input type="hidden" name="page" value="">
<input type="hidden" name="eventidx" value="<%= eventidx %>">
<input type="hidden" name="searchvalue" value="<%= searchvalue %>">
</form>


<script language="javascript" type="text/javascript">
function FnPageMove(page){
	document.pageFrm.page.value=page;
	document.pageFrm.submit();

}

function CheckSearch(){
	if (document.searchFrm.checkRight.checked) {
		document.searchFrm.searchInput.readOnly=false;
	} else {
		document.searchFrm.checkRight.checked=false;
		document.searchFrm.searchInput.value ='';
		document.searchFrm.searchInput.readOnly =true;
	}
}
function FnSearch(){
	var rightans = document.searchFrm.searchInput.value;
	document.location.href="/admin/sitemaster/diary_collection_2007/diary_quiz_event_entryList.asp?eventidx=<%= eventidx %>&searchvalue=" + rightans ;
}

function DoWinner(userid) {
	if(navigator.userAgent.indexOf("MSIE") !=-1) {
  	parent.winnerFrm.winnerList.value=parent.winnerFrm.winnerList.value + '\n' + userid;
 	}else {

		top.document.winnerFrm.winnerList.value=top.document.winnerFrm.winnerList.value + '\n' + userid;
	}
}

CheckSearch();
</script>







<!-- #include virtual="/lib/db/db2close.asp" -->

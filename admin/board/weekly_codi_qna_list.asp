<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/weekly_codi_qna_cls.asp" -->

<%

public Sub GetWeekelyList(byval Intv)
	dim sql
	response.write "<select name='masteridx' onchange=javascript:Selectidx('this.value');>--전체 선택 --</option>"
		sql = "select masteridx From [db_cs].[dbo].tbl_weekly_qna" + vbcrlf
		sql = sql + " where isusing='Y'" + vbcrlf
		sql = sql + " group by masteridx" + vbcrlf
		sql = sql + " order by masteridx desc" + vbcrlf
		
		rsget.open sql,dbget,1
		response.write "<option value=''>--선택--</option>"
		response.write "<option value=''>--전체보기--</option>"
	do until rsget.eof 
		response.write "<option value='" & rsget("masteridx") & "'>" & rsget("masteridx") & "</option>"
	rsget.movenext
	loop
	rsget.close
end Sub
	
dim masteridx,i,ix,page
masteridx=request("masteridx")
page=request("page")
if page="" then page=1
dim wqna
set wqna = new WeeklyQna
wqna.FCurrpage=page
wqna.FPageSize=20
wqna.getQnaList masteridx
%>
<script>
function MovePage(page,masteridx){
	document.ffrm.page.value=page;
	document.ffrm.masteridx.value=masteridx;
	document.ffrm.submit();
}
function Selectidx(v){
	
	document.ffrm.submit();
}
</script>
<table width="760" border="0" cellpadding="0" cellspacing="0" class="a" bgcolor="#FFFFFF">
	<form name="ffrm" action="/admin/board/weekly_codi_qna_list.asp" method="post">
	<input type="hidden" name="page" value="<%= page %>">
	<tr>
		<td><% GetWeekelyList masteridx%></td>
	</tr>
	</form>
<table> 
<table width="760" border="0" cellpadding="3" cellspacing="1" class="a" bgcolor="#000000">
	<% if masteridx<>"" then %>
	<tr bgcolor="FFFFFF">
		<td colspan="5"><a href="http://www.10x10.co.kr/guidebook/weekly_codi.asp?idx=<%= Masteridx %>" target="_blank">위클리 코디네이터 보기</a></td></td>
	</tr>
	<% end if %>
	<tr bgcolor="#DDDDFF">
		<td width="90" align="center">User Id</td>
		<td width="80" align="center">본문</td>
		<td align="center">질문 내용</td>
		<td width="70" align="center">답변 유무</td>
		<td width="100" align="center">등록 날짜</td>
	</tr>
	<% if wqna.FResultcount>0 then %>
	<% for i=0 to wqna.FResultCount -1 %>
	<tr bgcolor="FFFFFF">
		<td width="90" align="center"><%= wqna.FItemList(i).FUserid %></td>
		<td width="80" align="center"><%= wqna.FItemList(i).FMasteridx %></td>
		<td><a href="/admin/board/weekly_codi_qna_view.asp?idx=<%=wqna.FItemList(i).Fidx %>&masteridx=<%= Masteridx %>&page=<%= page %>"><%= Left(wqna.FItemList(i).FQuestion,35) %></a></td>
		<td width="70" align="center"><%= wqna.FItemList(i).FAnswerYN %></td>
		<td width="100" align="center"><%= left(wqna.FItemList(i).FRegdate,10) %></td>
	</tr>
	<% next %>
	<% end if %>
</table>
<table width="760" cellspacing="1" class="a" bgcolor=#3d3d3d>
  <tr bgcolor="#FFFFFF">
    <td align="center">
		<% if wqna.HasPreScroll then %>
			<a href="javascript:MovePage('<%= wqna.StartScrollPage-1 %>','<%=masteridx%>');">[prev]</a>
		<% else %>
			[prev]
		<% end if %>

		<% for ix=0 + wqna.StartScrollPage to wqna.FScrollCount + wqna.StartScrollPage - 1 %>
			<% if ix>wqna.FTotalpage then Exit for %>
				<% if CStr(page)=CStr(ix) then %>
					<font color="red">[<%= ix%>]</font>
				<% else %>
					<a href="javascript:MovePage('<%=ix%>','<%=masteridx%>');">[<%= ix %>]</a>
				<% end if %>
		<% next %>

		<% if wqna.HasNextScroll then %>
			<a href="javascript:MovePage('<%=ix%>','<%=masteridx%>');">[next]</a>
		<% else %>
			[next]
		<% end if %>
	</td>
  </tr>
</table>

<% set wqna=nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
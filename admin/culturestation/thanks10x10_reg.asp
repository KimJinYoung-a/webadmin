<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : Culture Station Thanks 10x10 등록  
' History : 2009.04.02 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/culturestation/culturestation_class.asp"-->

<% 
dim idx,idxsum ,itemid, rectitemid,mode ,gubun
dim sql , oip,i , page
	idx = request("idx")
	mode = request("mode")
	gubun = request("gubun")
	if page = "" then page = 1

set oip = new cthanks10x10_list
	oip.FPageSize = 20
	oip.FCurrPage = page
	oip.frectidx = idx
	oip.fthanks10x10_list()
%>

<script language="javascript">

function reg(){
frm.submit();
}

function commit_del(){
var ret;
	ret = confirm('답변을 삭제 하시겠습니까?');
	
	if(ret){
		frm.mode.value = 'comment_del'
		frm.submit();
	}
}

</script>
	
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="post" action="/admin/culturestation/thanks10x10_process.asp">
	<input type="hidden" name="mode" value="<%= gubun %>">
	<input type="hidden" name="idx" value="<%= idx %>">
	<tr bgcolor="#FFFFFF">
		<td align="center" colspan=2 bgcolor="<%= adminColor("tabletop") %>"><b>고객글</b><br></td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td align="center" width=50>
			<%= drawgubun(oip.fitemlist(0).fgubun) %>
		</td>
		<td align="center">
			<textarea rows=15 cols=100><%= oip.fitemlist(0).fcontents %></textarea>
		</td>
	</tr>			
</table><br>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr bgcolor="#FFFFFF">
		<td align="center" bgcolor="<%= adminColor("tabletop") %>"><b>답변하기</b>
		</td>
	</tr>	
	<tr bgcolor="#FFFFFF">
		<td align="center"><textarea rows=10 cols=100 name="comment"><%= oip.fitemlist(0).fcomment %></textarea>
		</td>
	</tr>			
	<tr bgcolor="#FFFFFF">
		<td align="center">
			<% if gubun = "add" then %>
				<input type="button" class="button" value="저장" onclick="javascript:reg();">	
			<% elseif gubun = "edit" then %>
				<input type="button" class="button" value="수정" onclick="javascript:reg();">
				<input type="button" class="button" value="답변삭제" onclick="javascript:commit_del();">			
			<% end if %>
		</td>
	</tr>	
</table>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->

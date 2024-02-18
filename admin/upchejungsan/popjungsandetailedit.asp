<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->

<%
dim id , gubun, itemid, itemoption, sellcash, suplycash
dim itemname, itemoptionname
dim mode , edit_sellcash , edit_suplycash

mode	= request.form("mode")
id		= request.form("id")
gubun	= request.form("gubun")
itemid	= request.form("itemid")
itemoption	= request.form("itemoption")
sellcash	= request.form("sellcash")
suplycash	= request.form("suplycash")
itemname = request.form("itemname") ''html2db()
itemoptionname = request.form("itemoptionname") ''html2db()

itemname        = replace(itemname,"||39||","'")
itemoptionname  = replace(itemoptionname,"||39||","'")


edit_sellcash	= request.form("edit_sellcash")
edit_suplycash	= request.form("edit_suplycash")

'response.write itemoptionname

dim sqlStr
dim currfinishflag
dim AssignedRows
if mode="edit" then

	sqlStr = "select top 1 id, finishflag from [db_jungsan].[dbo].tbl_designer_jungsan_master"
	sqlStr = sqlStr + " where id=" + id + ""
	rsget.Open sqlStr,dbget,1
	if Not rsget.Eof then
		currfinishflag = rsget("finishflag")
	end if
	rsget.Close

	if Not ((currfinishflag="0") or (currfinishflag="1")) then
		response.write "<script language=javascript>"
		response.write "alert('현재 수정중 또는 업체 확인대기 상태가 아닙니다. 수정 하실 수 없습니다.');"
		response.write "</script>"
	else
		sqlStr = "update [db_jungsan].[dbo].tbl_designer_jungsan_detail" + VbCrlf
		sqlStr = sqlStr + " set sellcash=" + CStr(edit_sellcash) + VbCrlf
		sqlStr = sqlStr + " ,suplycash=" + CStr(edit_suplycash) + VbCrlf
		sqlStr = sqlStr + " where masteridx=" + CStr(id) + VbCrlf
		sqlStr = sqlStr + " and gubuncd='" + gubun + "'" + VbCrlf
		sqlStr = sqlStr + " and itemid=" + CStr(itemid) + VbCrlf
		sqlStr = sqlStr + " and itemoption='" + CStr(itemoption) + "'" + VbCrlf
		sqlStr = sqlStr + " and sellcash= " + CStr(sellcash) + VbCrlf
		sqlStr = sqlStr + " and suplycash= " + CStr(suplycash) + VbCrlf
		sqlStr = sqlStr + " and replace(itemname,'&amp;','&')='" + html2db(itemname) + "'" + VbCrlf
		sqlStr = sqlStr + " and itemoptionname='" + html2db(itemoptionname) + "'" + VbCrlf

		dbget.Execute sqlStr,AssignedRows
        
        response.write "적용갯수=" & AssignedRows

		if gubun="upche" then
			sqlStr = "update [db_jungsan].[dbo].tbl_designer_jungsan_master" + VbCrlf
			sqlStr = sqlStr + " set ub_cnt=T.cnt" + VbCrlf
			sqlStr = sqlStr + " ,ub_totalsellcash=T.totalsellcash" + VbCrlf
			sqlStr = sqlStr + " ,ub_totalsuplycash=T.totalsuplycash" + VbCrlf
			sqlStr = sqlStr + " from (select count(d.mastercode) as cnt, sum(d.itemno*d.sellcash) as totalsellcash,sum(d.itemno*d.suplycash) as totalsuplycash" + VbCrlf
			sqlStr = sqlStr + " from [db_jungsan].[dbo].tbl_designer_jungsan_detail d" + VbCrlf
			sqlStr = sqlStr + " where masteridx=" + CStr(id) + VbCrlf
			sqlStr = sqlStr + " and gubuncd='upche') as T" + VbCrlf
			sqlStr = sqlStr + " where id=" + CStr(id)
			rsget.Open sqlStr,dbget,1
		elseif gubun="maeip" then
			sqlStr = "update [db_jungsan].[dbo].tbl_designer_jungsan_master" + VbCrlf
			sqlStr = sqlStr + " set me_cnt=T.cnt" + VbCrlf
			sqlStr = sqlStr + " ,me_totalsellcash=T.totalsellcash" + VbCrlf
			sqlStr = sqlStr + " ,me_totalsuplycash=T.totalsuplycash" + VbCrlf
			sqlStr = sqlStr + " from (select count(d.mastercode) as cnt, sum(d.itemno*d.sellcash) as totalsellcash,sum(d.itemno*d.suplycash) as totalsuplycash" + VbCrlf
			sqlStr = sqlStr + " from [db_jungsan].[dbo].tbl_designer_jungsan_detail d" + VbCrlf
			sqlStr = sqlStr + " where masteridx=" + CStr(id) + VbCrlf
			sqlStr = sqlStr + " and gubuncd='maeip') as T" + VbCrlf
			sqlStr = sqlStr + " where id=" + CStr(id)
			'response.write sqlStr
			rsget.Open sqlStr,dbget,1
		elseif gubun="witaksell" then
			sqlStr = "update [db_jungsan].[dbo].tbl_designer_jungsan_master" + VbCrlf
			sqlStr = sqlStr + " set wi_cnt=T.cnt" + VbCrlf
			sqlStr = sqlStr + " ,wi_totalsellcash=T.totalsellcash" + VbCrlf
			sqlStr = sqlStr + " ,wi_totalsuplycash=T.totalsuplycash" + VbCrlf
			sqlStr = sqlStr + " from (select count(d.mastercode) as cnt, sum(d.itemno*d.sellcash) as totalsellcash,sum(d.itemno*d.suplycash) as totalsuplycash" + VbCrlf
			sqlStr = sqlStr + " from [db_jungsan].[dbo].tbl_designer_jungsan_detail d" + VbCrlf
			sqlStr = sqlStr + " where masteridx=" + CStr(id) + VbCrlf
			sqlStr = sqlStr + " and gubuncd='witaksell') as T" + VbCrlf
			sqlStr = sqlStr + " where id=" + CStr(id)
			'response.write sqlStr
			rsget.Open sqlStr,dbget,1
		elseif gubun="witakchulgo" then
			sqlStr = "update [db_jungsan].[dbo].tbl_designer_jungsan_master" + VbCrlf
			sqlStr = sqlStr + " set et_cnt=T.cnt" + VbCrlf
			sqlStr = sqlStr + " ,et_totalsellcash=T.totalsellcash" + VbCrlf
			sqlStr = sqlStr + " ,et_totalsuplycash=T.totalsuplycash" + VbCrlf
			sqlStr = sqlStr + " from (select count(d.mastercode) as cnt, sum(d.itemno*d.sellcash) as totalsellcash,sum(d.itemno*d.suplycash) as totalsuplycash" + VbCrlf
			sqlStr = sqlStr + " from [db_jungsan].[dbo].tbl_designer_jungsan_detail d" + VbCrlf
			sqlStr = sqlStr + " where masteridx=" + CStr(id) + VbCrlf
			sqlStr = sqlStr + " and gubuncd='witakchulgo') as T" + VbCrlf
			sqlStr = sqlStr + " where id=" + CStr(id)
			'response.write sqlStr
			rsget.Open sqlStr,dbget,1
		elseif gubun="witakoffshop" then
			sqlStr = "update [db_jungsan].[dbo].tbl_designer_jungsan_master"
			sqlStr = sqlStr + " set sh_cnt=T.cnt"
			sqlStr = sqlStr + " ,sh_totalsellcash=T.totalsellcash"
			sqlStr = sqlStr + " ,sh_totalsuplycash=T.totalsuplycash"
			sqlStr = sqlStr + " from (select count(d.mastercode) as cnt, sum(d.itemno*d.sellcash) as totalsellcash,sum(d.itemno*d.suplycash) as totalsuplycash"
			sqlStr = sqlStr + " from [db_jungsan].[dbo].tbl_designer_jungsan_detail d"
			sqlStr = sqlStr + " where masteridx=" + CStr(id)
			sqlStr = sqlStr + " and gubuncd='witakoffshop') as T"
			sqlStr = sqlStr + " where id=" + CStr(id)
			'response.write sqlStr
			rsget.Open sqlStr,dbget,1
		end if
	end if
end if


dim itemno

sqlStr = "select d.itemid,d.itemoption,d.itemname,d.itemoptionname, sum(d.itemno) as itemno, d.sellcash, d.suplycash "
sqlStr = sqlStr + " from [db_jungsan].[dbo].tbl_designer_jungsan_detail d"
sqlStr = sqlStr + " where d.masteridx=" + CStr(id)
sqlStr = sqlStr + " and d.gubuncd='" + gubun + "'"
sqlStr = sqlStr + " and d.itemid=" + CStr(itemid)
sqlStr = sqlStr + " and d.itemoption='" + CStr(itemoption) + "'"
sqlStr = sqlStr + " and d.sellcash= " + CStr(sellcash)
sqlStr = sqlStr + " and d.suplycash= " + CStr(suplycash)
sqlStr = sqlStr + " and d.itemname='" + html2db(itemname) + "'"
sqlStr = sqlStr + " and d.itemoptionname='" + html2db(itemoptionname) + "'"
sqlStr = sqlStr + " group by d.itemid,d.itemoption,d.itemname,d.itemoptionname, d.sellcash, d.suplycash "

'response.write sqlStr

rsget.Open sqlStr,dbget,1
if not rsget.Eof then
	itemname		= rsget("itemname") ''db2html()
	itemoptionname	= rsget("itemoptionname") ''db2html()
	itemno 			= rsget("itemno")
end if
rsget.close
%>
<script language='javascript'>
function EditDetail(frm){
	if (frm.edit_sellcash.value.length<1){
		alert();
		frm.edit_sellcash.focus();
	}

	if (frm.edit_suplycash.value.length<1){
		alert();
		frm.edit_suplycash.focus();
	}

	if (confirm('일괄수정하시겠습니까?')){
		editform.submit();
	}
}
</script>
<table width="500" border="0" cellspacing="1" cellpadding="2" class="a" bgcolor=#3d3d3d>
<tr bgcolor="#DDDDFF" align="center">
	<td>상품코드</td>
	<td>옵션코드</td>
	<td>상품명</td>
	<td>옵션명</td>
	<td>갯수</td>
	<td>판매가</td>
	<td>매입가</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td><%= itemid %></td>
	<td><%= itemoption %></td>
	<td><%= itemname %></td>
	<td><%= itemoptionname %></td>
	<td align="center"><%= itemno %></td>
	<td align="right"><%= FormatNumber(sellcash,0) %></td>
	<td align="right"><%= FormatNumber(suplycash,0) %></td>
</tr>
<tr bgcolor="#FFFFFF">
	<form name="editform" method="post" action=''>
	<input type="hidden" name="mode" value='edit'>
	<input type="hidden" name="id" value='<%= id %>'>
	<input type="hidden" name="gubun" value='<%= gubun %>'>
	<input type="hidden" name="itemid" value='<%= itemid %>'>
	<input type="hidden" name="itemoption" value='<%= itemoption %>'>
	<input type="hidden" name="sellcash" value='<%= sellcash %>'>
	<input type="hidden" name="suplycash" value='<%= suplycash %>'>
	<input type="hidden" name="itemname" value="<%= replace(itemname,"'","'") %>">
	<input type="hidden" name="itemoptionname" value="<%= replace(itemoptionname,"'","'") %>">
	<td></td>
	<td></td>
	<td></td>
	<td></td>
	<td></td>
	<td align="right"><input type="text" name="edit_sellcash" value="<%= sellcash %>" size="5" maxlength="8"></td>
	<td align="right"><input type="text" name="edit_suplycash" value="<%= suplycash %>" size="5" maxlength="8"></td>
	</form>
</tr>
<tr bgcolor="#FFFFFF">
	<td colspan="7" align="center"><input type="button" value="일괄 수정" onclick="EditDetail(editform);"></td>
</tr>
</table>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
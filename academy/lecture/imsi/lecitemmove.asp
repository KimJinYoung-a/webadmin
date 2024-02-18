<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/academy/lib/academy_function.asp"-->

<script language='javascript'>
function moveLecItem(lecidx){
	if (confirm('이동 하시겠습니까?')){
		var popwin = window.open('lecmove.asp?mode=lecitem&lecidx=' + lecidx,'moveLecItem','width=400,height=400,scrollbars=yes,resizable=yes')
		popwin.focus();
	}
}

function delLecItem(lecidx){
	if (confirm('삭제 하시겠습니까?')){
		var popwin = window.open('lecmove.asp?mode=lecitemdel&lecidx=' + lecidx,'moveLecItem','width=400,height=400,scrollbars=yes,resizable=yes')
		popwin.focus();
	}
}

</script>
<%
dim sqlStr

'''링크 아이템 중복건수 체크 : 중복되는것중 사용안하는것 linkitemid -> -1로 설정
sqlStr = "select top 10 * from "
sqlStr = sqlStr + " (select linkitemid,count(idx) as cnt from "
sqlStr = sqlStr + " [db_contents].[dbo].tbl_lecture_item L,"
sqlStr = sqlStr + " [db_item].[dbo].tbl_item i"
sqlStr = sqlStr + " where L.linkitemid=i.itemid"
sqlStr = sqlStr + " group by linkitemid"
sqlStr = sqlStr + " ) T"
sqlStr = sqlStr + " where T.cnt>1"

rsget.Open sqlStr, dbget, 1

response.write "<br> <font color=blue>1..링크아이템 중복 강좌 체크.</font><br>"
do until rsget.eof
	response.write CStr(rsget("linkitemid")) + ":" + CStr(rsget("cnt")) + "건<br>"
	rsget.movenext
loop
rsget.close
%>

<%
''' 이전 안된 강좌
sqlStr = "select top 100 L.* from  "
sqlStr = sqlStr + " [db_contents].[dbo].tbl_lecture_item L "
sqlStr = sqlStr + " left join [db_academy].[dbo].tbl_lec_item N "
sqlStr = sqlStr + " on L.idx=N.idx "
sqlStr = sqlStr + " where N.idx is null "
''sqlStr = sqlStr + " and L.linkitemid<>-1 "
sqlStr = sqlStr + " order by L.idx"

response.write "<br> <font color=blue>2..이전 안된 강좌 체크.</font><br>"

rsget.Open sqlStr, dbget, 1
do until rsget.eof
%>
	<%= CStr(rsget("idx")) %>:<%= CStr(rsget("linkitemid"))  %>:<%= CStr(rsget("lecturerid"))  %>:<%= CStr(db2html(rsget("lectitle")))  %>:<%= CStr(rsget("mastercode"))  %>
	<input type="button" value="이전" onclick="moveLecItem('<%= CStr(rsget("idx")) %>')">
	<br>
<%
	rsget.movenext
loop
rsget.close
%>

<%
''' 이전 된 강좌중 내역이 다른것 - 기본정보
sqlStr = "select top 100 L.* from  "
sqlStr = sqlStr + " (select L.*, i.limitsold from "
sqlStr = sqlStr + " [db_contents].[dbo].tbl_lecture_item L, [db_item].[dbo].tbl_item i"
sqlStr = sqlStr + " where L.linkitemid=i.itemid"
sqlStr = sqlStr + " ) as L"
sqlStr = sqlStr + " , [db_academy].[dbo].tbl_lec_item N "
sqlStr = sqlStr + " where  L.idx=N.idx "
sqlStr = sqlStr + " and ("
sqlStr = sqlStr + " 		L.mastercode<>N.lec_date "
sqlStr = sqlStr + " 	or	L.lecturerid<>N.lecturer_id "
sqlStr = sqlStr + " 	or	L.lecsum<>N.lec_cost"
sqlStr = sqlStr + " 	or	L.matsum<>N.mat_cost"
sqlStr = sqlStr + " 	or	L.matinclude<>N.matinclude_yn"
sqlStr = sqlStr + " 	or	L.properperson<>N.limit_count"
sqlStr = sqlStr + " 	or	IsNULL(L.limitsold,0)<>N.limit_sold "
sqlStr = sqlStr + " 	or	L.lecdate01<>N.lec_startday1"
sqlStr = sqlStr + " 	or	L.lecdate02<>N.lec_startday2"
sqlStr = sqlStr + " 	or	L.lecdate03<>N.lec_startday3"
sqlStr = sqlStr + " 	or	L.lecdate04<>N.lec_startday4"
sqlStr = sqlStr + " 	or	L.lecdate05<>N.lec_startday5"
sqlStr = sqlStr + " 	or	L.leccount<>N.lec_count"
sqlStr = sqlStr + " 	or	L.lectime<>N.lec_time"
sqlStr = sqlStr + " 	or	L.lecperiod<>N.lec_period"
sqlStr = sqlStr + " 	or	L.lecspace<>N.lec_space"
sqlStr = sqlStr + " 	or	convert(varchar(255),L.leccontents)<>convert(varchar(255),N.lec_outline)"
sqlStr = sqlStr + " 	or	convert(varchar(255),L.leccurry)<>convert(varchar(255),N.lec_contents)"
sqlStr = sqlStr + " 	or	convert(varchar(255),L.lecetc)<>convert(varchar(255),N.lec_etccontents)"
sqlStr = sqlStr + " 	or	L.isusing<>N.isusing"
sqlStr = sqlStr + " 	or	L.isusing<>N.disp_yn"
sqlStr = sqlStr + " 	or	(case when L.regfinish='N' then 'Y' else 'N' end )<>N.reg_yn"

sqlStr = sqlStr + " ) "
sqlStr = sqlStr + " order by L.idx desc"


response.write "<br> <font color=blue>3..이전된 강좌 중 내용 변경 체크. - 1</font><br>"

rsget.Open sqlStr, dbget, 1
do until rsget.eof
%>
	<%= CStr(rsget("idx")) %>:<%= CStr(rsget("linkitemid"))  %>:<%= CStr(rsget("lecturerid"))  %>:<%= CStr(db2html(rsget("lectitle")))  %>:<%= CStr(rsget("mastercode"))  %>
	<input type="button" value="삭제" onclick="delLecItem('<%= CStr(rsget("idx")) %>')">
	<br>
<%
	rsget.movenext
loop
rsget.close
%>


<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
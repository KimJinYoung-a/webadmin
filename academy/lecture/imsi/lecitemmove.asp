<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/academy/lib/academy_function.asp"-->

<script language='javascript'>
function moveLecItem(lecidx){
	if (confirm('�̵� �Ͻðڽ��ϱ�?')){
		var popwin = window.open('lecmove.asp?mode=lecitem&lecidx=' + lecidx,'moveLecItem','width=400,height=400,scrollbars=yes,resizable=yes')
		popwin.focus();
	}
}

function delLecItem(lecidx){
	if (confirm('���� �Ͻðڽ��ϱ�?')){
		var popwin = window.open('lecmove.asp?mode=lecitemdel&lecidx=' + lecidx,'moveLecItem','width=400,height=400,scrollbars=yes,resizable=yes')
		popwin.focus();
	}
}

</script>
<%
dim sqlStr

'''��ũ ������ �ߺ��Ǽ� üũ : �ߺ��Ǵ°��� �����ϴ°� linkitemid -> -1�� ����
sqlStr = "select top 10 * from "
sqlStr = sqlStr + " (select linkitemid,count(idx) as cnt from "
sqlStr = sqlStr + " [db_contents].[dbo].tbl_lecture_item L,"
sqlStr = sqlStr + " [db_item].[dbo].tbl_item i"
sqlStr = sqlStr + " where L.linkitemid=i.itemid"
sqlStr = sqlStr + " group by linkitemid"
sqlStr = sqlStr + " ) T"
sqlStr = sqlStr + " where T.cnt>1"

rsget.Open sqlStr, dbget, 1

response.write "<br> <font color=blue>1..��ũ������ �ߺ� ���� üũ.</font><br>"
do until rsget.eof
	response.write CStr(rsget("linkitemid")) + ":" + CStr(rsget("cnt")) + "��<br>"
	rsget.movenext
loop
rsget.close
%>

<%
''' ���� �ȵ� ����
sqlStr = "select top 100 L.* from  "
sqlStr = sqlStr + " [db_contents].[dbo].tbl_lecture_item L "
sqlStr = sqlStr + " left join [db_academy].[dbo].tbl_lec_item N "
sqlStr = sqlStr + " on L.idx=N.idx "
sqlStr = sqlStr + " where N.idx is null "
''sqlStr = sqlStr + " and L.linkitemid<>-1 "
sqlStr = sqlStr + " order by L.idx"

response.write "<br> <font color=blue>2..���� �ȵ� ���� üũ.</font><br>"

rsget.Open sqlStr, dbget, 1
do until rsget.eof
%>
	<%= CStr(rsget("idx")) %>:<%= CStr(rsget("linkitemid"))  %>:<%= CStr(rsget("lecturerid"))  %>:<%= CStr(db2html(rsget("lectitle")))  %>:<%= CStr(rsget("mastercode"))  %>
	<input type="button" value="����" onclick="moveLecItem('<%= CStr(rsget("idx")) %>')">
	<br>
<%
	rsget.movenext
loop
rsget.close
%>

<%
''' ���� �� ������ ������ �ٸ��� - �⺻����
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


response.write "<br> <font color=blue>3..������ ���� �� ���� ���� üũ. - 1</font><br>"

rsget.Open sqlStr, dbget, 1
do until rsget.eof
%>
	<%= CStr(rsget("idx")) %>:<%= CStr(rsget("linkitemid"))  %>:<%= CStr(rsget("lecturerid"))  %>:<%= CStr(db2html(rsget("lectitle")))  %>:<%= CStr(rsget("mastercode"))  %>
	<input type="button" value="����" onclick="delLecItem('<%= CStr(rsget("idx")) %>')">
	<br>
<%
	rsget.movenext
loop
rsget.close
%>


<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
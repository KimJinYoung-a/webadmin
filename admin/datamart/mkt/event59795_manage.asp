<%@ language=vbscript %>
<% option Explicit %>
<% Response.CharSet = "euc-kr" %>
<%
'###########################################################
' Description : 슈퍼백의 기적 통계
' History : 2015.03.16 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<%
dim eCode, subscriptcount, userid, evtTotalCnt, sqlStr , itemid
Dim returndate  : returndate = 	request("returndate")

If returndate = "" Then returndate = Date()

IF application("Svr_Info") = "Dev" THEN
	eCode 		= "21494"
	itemid		= getdateitem(returndate)
Else
	eCode 		= "59795"
	itemid		= getdateitem(returndate)
End If

If session("ssBctId")="winnie" Or session("ssBctId")="gawisonten10" Or session("ssBctId") = "edojun" Or session("ssBctId") = "tozzinet" Or session("ssBctId") = "thensi7" Then

Else

	response.write "<script>alert('관계자만 볼 수 있는 페이지 입니다.');window.close();</script>"
	response.End

End If

dim winnumber
winnumber=0

sqlStr = "select top 1 bigo as winnumber"
sqlStr = sqlStr & " from db_temp.dbo.tbl_event_etc_yongman"
sqlStr = sqlStr & " where event_code="& eCode &" and isusing='Y'"

'response.write sqlStr & "<br>"
rsget.Open sqlStr,dbget
if Not(rsget.EOF or rsget.BOF) Then
	winnumber=rsget("winnumber")
else
	winnumber=0
End If
rsget.close
winnumber=getNumeric(winnumber)
%>
			
<style type="text/css">
.table {width:900px; margin:0 auto; font-family:'malgun gothic'; border-collapse:collapse;}
.table th {padding:12px 0; font-size:13px; font-weight:bold;  color:#fff; background:#444;}
.table td {padding:12px 3px; font-size:12px; border:1px solid #ddd; border-bottom:2px solid #ddd;}
.table td.lt {text-align:left; padding:12px 10px;}
.tBtn {display:inline-block; border:1px solid #2b90b6; background:#03a0db; padding:0 10px 2px; line-height:26px; height:26px; font-weight:bold; border-radius:5px; color:#fff !important;}
.tBtn:hover {text-decoration:none;}
</style>
<script type="text/javascript">

function regwinnumber(){
	if(winnumberfrm.winnumber.value!=''){
		if (!IsDouble(winnumberfrm.winnumber.value)){
			alert('확률은 숫자만 가능합니다.');
			winnumberfrm.winnumber.focus();
			return;
		}
		
		winnumberfrm.action="/admin/datamart/mkt/event59795_manage_process.asp"
		winnumberfrm.mode.value="winnumber";
		winnumberfrm.target="evtFrmProc";
		winnumberfrm.submit();
		return;
	}else{
		alert('확률을 입력해 주세요');
		winnumberfrm.winnumber.focus();
		return;
	}
	
}

</script>
</head>
<body>

<table style="margin:0 auto;text-align:center;">
<tr>
	<td>
		<strong><font size=5>슈퍼백의 기적 관리자<br>현재 보고 있는 날짜 : <%=returndate%>일</font></strong>
		<br><span style="color:red;">부하가 가는 페이지입니다. 막 누르지 마세요.</span>
	</td>
</tr>
</table>

<table style="margin:0 auto;text-align:center;margin-top:50px;">
<tr>
	<td>
		<span style="color:red;">
			1부터 - 1000 사이 비율로 입력해 주시면 됩니다.
			<br>0.1% -> 2 입력
			<br>1% -> 11 입력
			<br>10% -> 101 입력
			<br>50% -> 501 입력
			<br>90% -> 901 입력
		</span>
	</td>
</tr>
</table>
<table class="table" style="width:90%;">
<colgroup>
	<col width="*" />
</colgroup>
<tr align="center" bgcolor="#E6E6E6">
	<th><strong>확률조정 ※ 실시간</strong></th>
</tr>
<tr bgcolor="#FFFFFF" align="center">
	<td>
		<form name="winnumberfrm" method="post" action="" style="margin:0px;">
		<input type="hidden" name="mode">
		<input type="hidden" name="evt_code" value="<%= eCode %>">
		<input type="text" name="winnumber" value="<%= winnumber %>" size=3 maxlength=4>
		</form>	
		<input type="button" onclick="regwinnumber();" value="저장">
	</td>
</tr>
</table>

<table style="margin:0 auto;text-align:center;margin-top:50px;">
<tr>
	<td>
		날짜를 클릭하시면 하단에 통계치가 노출 됩니다.
	</td>
</tr>
</table>
<table class="table" style="width:90%;">
<colgroup>
	<col width="*" />
	<col width="*" />
	<col width="*" />
	<col width="*" />
	<col width="*" />
	<col width="*" />
</colgroup>
<tr align="center" bgcolor="#E6E6E6">
	<th colspan="8"><strong>날짜</strong></th>
</tr>
<tr bgcolor="#FFFFFF" align="center">
	<td><a href="http://webadmin.10x10.co.kr/admin/datamart/mkt/event59795_manage.asp?returndate=2015-03-16">2015-03-16</a></td>
	<td><a href="http://webadmin.10x10.co.kr/admin/datamart/mkt/event59795_manage.asp?returndate=2015-03-17">2015-03-17</a></td>
	<td><a href="http://webadmin.10x10.co.kr/admin/datamart/mkt/event59795_manage.asp?returndate=2015-03-18">2015-03-18</a></td>
	<td><a href="http://webadmin.10x10.co.kr/admin/datamart/mkt/event59795_manage.asp?returndate=2015-03-19">2015-03-19</a></td>
	<td><a href="http://webadmin.10x10.co.kr/admin/datamart/mkt/event59795_manage.asp?returndate=2015-03-20">2015-03-20</a></td>
	<td bgcolor="skyblue"><a href="http://webadmin.10x10.co.kr/admin/datamart/mkt/event59795_manage.asp?returndate=2015-03-21">2015-03-21</a></td>
</tr>																				            
<tr bgcolor="#FFFFFF" align="center">												            
	<td bgcolor="skyblue"><a href="http://webadmin.10x10.co.kr/admin/datamart/mkt/event59795_manage.asp?returndate=2015-03-22">2015-03-22</a></td>
	<td bgcolor="<%=chkiif(Date()="2015-03-23","FFFF09","")%>"><a href="http://webadmin.10x10.co.kr/admin/datamart/mkt/event59795_manage.asp?returndate=2015-03-23">2015-03-23</a></td>
	<td bgcolor="<%=chkiif(Date()="2015-03-24","FFFF09","")%>"><a href="http://webadmin.10x10.co.kr/admin/datamart/mkt/event59795_manage.asp?returndate=2015-03-24">2015-03-24</a></td>
	<td bgcolor="<%=chkiif(Date()="2015-03-25","FFFF09","")%>"><a href="http://webadmin.10x10.co.kr/admin/datamart/mkt/event59795_manage.asp?returndate=2015-03-25">2015-03-25</a></td>
</tr>
</table>

<table style="margin:0 auto;text-align:center;margin-top:30px;">
<tr>
	<td>시간대 별 클릭수 ※ 약 30분 정도의 오차가 있습니다.</td>
</tr>
</table>
<table class="table" style="width:90%;">
<colgroup>
	<col width="12.5%" />
	<col width="12.5%" />
	<col width="12.5%" />
	<col width="12.5%" />
	<col width="12.5%" />
	<col width="12.5%" />
	<col width="12.5%" />
	<col width="12.5%" />
</colgroup>
<tr align="center" bgcolor="#E6E6E6">
	<th><strong>시간</strong></th>
	<th><strong>M-Main 클릭수</strong></th>
	<th><strong>A-Main 클릭수</strong></th>
	<th><strong>kakao 클릭수</strong></th>
	<th><strong>회원가입수</strong></th>
	<th><strong>상품판매수</strong></th>
</tr>
</table>
<table class="table" style="width:90%;" border="0" cellspacing="0" cellpadding="0">
<tr>
	<td width="148">
		<table width="100%">
		<% '// 시간
			Dim i1
			For i1 = 0 To 23
		%>
		<tr style="text-align:center;">
			<td bgcolor="<%=chkiif(CInt(i1)=hour(now),"FFFF09","")%>"><%=i1%></td>
		</tr>
		<%
			Next
		%>
		</table>
	</td>
	<td>
		<table width="100%">
		<%
			sqlStr = " ; WITH [Hours] ([Hour]) AS "&_
					" (  "&_
					" SELECT TOP 24 ROW_NUMBER() OVER (ORDER BY [object_id]) AS [Hour]  "&_
					" FROM sys.objects  "&_
					" ORDER BY [object_id]  "&_
					" )  "&_
					" 	SELECT h.[Hour]-1 as Hour  "&_
					" 		,isnull(l.[mo_main],0) as mo_main  "&_
					" 		,isnull(l.[app_main],0) as app_main  "&_
					" 		,isnull(l.[kakao],0) as kakao  "&_
					" 	FROM [Hours] h  "&_
					" 	LEFT OUTER JOIN ( "&_
					" 		select   "&_
					" 			datepart(hh,regdate) as hhtime   "&_
					" 			, count(case when chkid = 'mo_main' then chkid end) as mo_main  "&_
					" 			, count(case when chkid = 'app_main' then chkid end) as app_main  "&_
					" 			, count(case when chkid = 'kakao' then chkid end) as kakao  "&_
					" 		from db_datamart.dbo.tbl_event_click_log   "&_
					" 		where eventid='"& eCode &"'   "&_
					" 		and convert(varchar(10),regdate,120) between '"& returndate &"' and '"& returndate &"'  "&_
					" 		group by datepart(hh,regdate) "&_
					" 	) as l "&_
					" ON h.[Hour]-1 = l.[hhtime] " &_
					" order by h.[Hour] asc "
			
			'Response.write sqlStr

			db3_rsget.Open sqlStr,db3_dbget
			Dim mo_main , app_main , kakao
			mo_main = 0
			app_main = 0
			kakao = 0

			if Not(db3_rsget.EOF or db3_rsget.BOF) Then
				Do Until db3_rsget.eof
		%>
		<tr style="text-align:center;">
			<td width="15%" bgcolor="<%=chkiif(db3_rsget("mo_main")>0,"skyblue","")%>"><%=db3_rsget("mo_main")%></td>
			<td width="15%" bgcolor="<%=chkiif(db3_rsget("app_main")>0,"skyblue","")%>"><%=db3_rsget("app_main")%></td>
			<td width="15%" bgcolor="<%=chkiif(db3_rsget("kakao")>0,"skyblue","")%>"><%=db3_rsget("kakao")%></td>
		</tr>
		<%
				mo_main = mo_main + db3_rsget("mo_main")
				app_main = app_main + db3_rsget("app_main")
				kakao = kakao + db3_rsget("kakao")

				db3_rsget.movenext
				Loop
			End If
			db3_rsget.close
		%>
		</table>
	</td>
	<td width="148">
		<table width="100%">
		<%
			sqlStr = " ; WITH [Hours] ([Hour]) AS "&_
					" (  "&_
					" SELECT TOP 24 ROW_NUMBER() OVER (ORDER BY [object_id]) AS [Hour]  "&_
					" FROM sys.objects  "&_
					" ORDER BY [object_id]  "&_
					" )  "&_
					" 	SELECT h.[Hour]-1 as Hour  "&_
					" 		,isnull(n.[usercnt],0) as usercnt "&_
					" 	FROM [Hours] h  "&_
					" 	LEFT OUTER JOIN ( "&_
					" 	select  "&_
					" 		count(*) as usercnt   "&_
					" 		, datepart(hh,n.regdate) as nhhdate   "&_
					" 		, convert(varchar(10),n.regdate,120) as ndate   "&_
					" 	from db_user.dbo.tbl_user_n as n  "&_
					" 	where left(n.eventid,10)= 'mobile_app'  "&_
					" 	and convert(varchar(10),n.regdate,120)  between '"& returndate &"' and '"& returndate &"'  "&_
					" 	group by datepart(hh,n.regdate) , convert(varchar(10),n.regdate,120)  "&_
					" 	) as n  "&_
					" ON h.[Hour]-1 = n.[nhhdate]  " &_
					" order by h.[Hour] asc "

			
			'Response.write sqlStr
			
			db3_rsget.Open sqlStr,db3_dbget
			Dim usercnt : usercnt = 0
			if Not(db3_rsget.EOF or db3_rsget.BOF) Then
				Do Until db3_rsget.eof
		%>
		<tr style="text-align:center;">
			<td width="100%" bgcolor="<%=chkiif(db3_rsget("usercnt")>0,"skyblue","")%>"><%=db3_rsget("usercnt")%></td>
		</tr>
		<%
				usercnt = usercnt + db3_rsget("usercnt")
				db3_rsget.movenext
				Loop
			End If
			db3_rsget.close
		%>
		</table>
	</td>
	<td width="148">
		<table width="100%">
		<%
			sqlStr = " ; WITH [Hours] ([Hour]) AS "&_
					" (  "&_
					" SELECT TOP 24 ROW_NUMBER() OVER (ORDER BY [object_id]) AS [Hour]  "&_
					" FROM sys.objects  "&_
					" ORDER BY [object_id]  "&_
					" )  "&_
					" 	SELECT h.[Hour]-1 as Hour  "&_
					" 		,isnull(m.[ordercnt],0) as ordercnt "&_
					" 	FROM [Hours] h  "&_
					" 	LEFT OUTER JOIN ( "&_
					" 		select   "&_
					" 			count(*) as ordercnt   "&_
					" 			, datepart(hh,m.regdate) as orderhhdate   "&_
					" 			, convert(varchar(10),m.regdate,120) as orderdate   "&_
					" 		from db_order.dbo.tbl_order_master as m  "&_
					" 		inner join db_order.dbo.tbl_order_detail as d   "&_
					" 		on m.orderserial=d.orderserial   "&_
					" 		where m.jumundiv<>'9' and m.ipkumdiv > 3   "&_
					" 		and m.cancelyn = 'N' and d.cancelyn<>'Y' and d.itemid<>'0' and d.itemid = '"& itemid &"'   "&_
					" 		and convert(varchar(10),m.regdate,120) between '"& returndate &"' and '"& returndate &"'  "&_
					" 		group by datepart(hh,m.regdate) , convert(varchar(10),m.regdate,120) "&_
					" 	) as m "&_
					" ON h.[Hour]-1 = m.[orderhhdate] " &_
					" order by h.[Hour] asc "

			'Response.write sqlStr

			db3_rsget.Open sqlStr,db3_dbget
			Dim ordercnt : ordercnt = 0
			if Not(db3_rsget.EOF or db3_rsget.BOF) Then
				Do Until db3_rsget.eof
		%>
		<tr style="text-align:center;">
			<td width="100%" bgcolor="<%=chkiif(db3_rsget("ordercnt")>0,"skyblue","")%>"><%=db3_rsget("ordercnt")%></td>
		</tr>
		<%
				ordercnt = ordercnt + db3_rsget("ordercnt")
				db3_rsget.movenext
				Loop
			End If
			db3_rsget.close
		%>
		</table>
	</td>
</tr>
</table>
<table class="table" style="width:90%; text-align:center;" border="0" cellspacing="0" cellpadding="0">
<tr>
	<td width="12.5%">합계</td>
	<td width="12.5%"><%=mo_main%></td>
	<td width="12.5%"><%=app_main%></td>
	<td width="12.5%"><%=kakao%></td>
	<td width="12.5%"><%=usercnt%></td>
	<td width="12.5%"><%=ordercnt%></td>
</tr>
</table>

<% If session("ssBctId")="gawisonten10" or session("ssBctId") = "motions" Or session("ssBctId") = "greenteenz" Or session("ssBctId") = "thensi7" Or session("ssBctId") = "tozzinet" Then %>
	<table style="margin:0 auto;text-align:center;margin-top:30px;">
		<tr>
			<td>구매자 선착순 1000명 ※ 약 30분 정도의 오차가 있습니다.</td>
		</tr>
	</table>
	<table class="table" style="width:90%;">
	<colgroup>
		<col width="5%" />
		<col width="10%" />
		<col width="*" />
		<col width="*" />
	</colgroup>
	<tr align="center" bgcolor="#E6E6E6">
		<th><strong>구매순위</strong></th>
		<th><strong>주문번호</strong></th>
		<th><strong>회원아이디</strong></th>
		<th><strong>주문시간대</strong></th>
	</tr>
	<%
		dim tmpval
		tmpval = 0
		sqlStr = " select top 1000 m.orderserial , m.userid , m.regdate from db_order.dbo.tbl_order_master as m "&_
				" inner join db_order.dbo.tbl_order_detail as d "&_
				" on m.orderserial=d.orderserial "&_
				" where m.jumundiv<>'9' and m.ipkumdiv > 3 and m.cancelyn = 'N'  "&_
				" and d.cancelyn<>'Y' and d.itemid<>'0' and d.itemid = '"& itemid &"' "&_
				" and m.regdate between '"&returndate&" 10:00:00' and '"&returndate&" 23:59:59' "&_
				" order by m.orderserial asc"

		db3_rsget.Open sqlStr,dbget
		if Not(db3_rsget.EOF or db3_rsget.BOF) Then
			Do Until db3_rsget.eof
			tmpval = tmpval + 1
	%>
	<tr bgcolor="#FFFFFF" align="center">
		<td><%= tmpval %></td>
		<td><%=db3_rsget("orderserial")%></a></td>
		<td><%=db3_rsget("userid")%></td>
		<td><%=db3_rsget("regdate")%></td>
	</tr>
	<%
			db3_rsget.movenext
			Loop
		End If
		db3_rsget.close
	%>
	</table>
<% End If %>

<iframe id="evtFrmProc" name="evtFrmProc" src="about:blank" frameborder="0" width=0 height=0></iframe>

<%
function getdateitem(dateval)
	dim tmpdateitem

	if dateval="" then
		getdateitem=""
		exit function
	end if

	IF application("Svr_Info") = "Dev" THEN
		if dateval="2015-03-16" then
			tmpdateitem=662881
		elseif dateval="2015-03-17" then
			tmpdateitem=662880
		elseif dateval="2015-03-18" then
			tmpdateitem=662878
		elseif dateval="2015-03-19" then
			tmpdateitem=662872
		elseif dateval="2015-03-20" then
			tmpdateitem=662867
		elseif dateval="2015-03-23" then
			tmpdateitem=662865
		elseif dateval="2015-03-24" then
			tmpdateitem=662842
		elseif dateval="2015-03-25" then
			tmpdateitem=662838
		else
			tmpdateitem=""
		end if
	Else
		if dateval="2015-03-16" then
			tmpdateitem=1224460
		elseif dateval="2015-03-17" then
			tmpdateitem=1224461
		elseif dateval="2015-03-18" then
			tmpdateitem=1224462
		elseif dateval="2015-03-19" then
			tmpdateitem=1224466
		elseif dateval="2015-03-20" then
			tmpdateitem=1224472
		elseif dateval="2015-03-23" then
			tmpdateitem=1224478
		elseif dateval="2015-03-24" then
			tmpdateitem=1224479
		elseif dateval="2015-03-25" then
			tmpdateitem=1224481
		else
			tmpdateitem=""
		end if
	End If

	getdateitem=tmpdateitem
end function
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->
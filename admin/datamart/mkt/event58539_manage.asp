<%@ language=vbscript %>
<% option Explicit %>
<% Response.CharSet = "euc-kr" %>
<%
'###########################################################
' Description : 58539 어벤저박스 관리자
' History : 2015-01-12 이종화 생성
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
	eCode 		= "21435" '// 응모 이벤트 코드 앱
	itemid		= "1000095"
Else
	eCode 		= "58539" '// 응모 이벤트 코드
	itemid		= "1197448" '//어벤져박스 상품
End If
'winnie','gawisonten10','greenteenz','edojun'

If session("ssBctId")="winnie" Or session("ssBctId")="gawisonten10" Or session("ssBctId") ="greenteenz" Or session("ssBctId") = "edojun" Or session("ssBctId") = "tozzinet" Or session("ssBctId") = "motions" Then

Else

	response.write "<script>alert('관계자만 볼 수 있는 페이지 입니다.');window.close();</script>"
	response.End

End If
%>
<style type="text/css">
.table {width:900px; margin:0 auto; font-family:'malgun gothic'; border-collapse:collapse;}
.table th {padding:12px 0; font-size:13px; font-weight:bold;  color:#fff; background:#444;}
.table td {padding:12px 3px; font-size:12px; border:1px solid #ddd; border-bottom:2px solid #ddd;}
.table td.lt {text-align:left; padding:12px 10px;}
.tBtn {display:inline-block; border:1px solid #2b90b6; background:#03a0db; padding:0 10px 2px; line-height:26px; height:26px; font-weight:bold; border-radius:5px; color:#fff !important;}
.tBtn:hover {text-decoration:none;}
.table td input {border:1px solid #ddd; height:30px; padding:0 3px; font-size:14px; color:#ec0d02; text-align:right;}
</style>
</head>
<body>
<p>&nbsp;</p>

<table style="margin:0 auto;text-align:center;">
	<tr>
		<td><strong>어벤져박스 통계<br><span style="color:red;">날짜를 누르면 해당 날짜의 데이터를 볼 수 있습니다.</span></strong><br/>※ 약 30분정도의 오차가 있을 수 있습니다.</td>
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
		<td><a href="http://webadmin.10x10.co.kr/admin/datamart/mkt/event58539_manage.asp?returndate=2015-01-19">2015-01-19</a></td>
		<td><a href="http://webadmin.10x10.co.kr/admin/datamart/mkt/event58539_manage.asp?returndate=2015-01-20">2015-01-20</a></td>
		<td><a href="http://webadmin.10x10.co.kr/admin/datamart/mkt/event58539_manage.asp?returndate=2015-01-21">2015-01-21</a></td>
		<td><a href="http://webadmin.10x10.co.kr/admin/datamart/mkt/event58539_manage.asp?returndate=2015-01-22">2015-01-22</a></td>
		<td><a href="http://webadmin.10x10.co.kr/admin/datamart/mkt/event58539_manage.asp?returndate=2015-01-23">2015-01-23</a></td>
		<td bgcolor="skyblue"><a href="http://webadmin.10x10.co.kr/admin/datamart/mkt/event58539_manage.asp?returndate=2015-01-24">2015-01-24</a></td>
	</tr>																				            
	<tr bgcolor="#FFFFFF" align="center">												            
		<td bgcolor="skyblue"><a href="http://webadmin.10x10.co.kr/admin/datamart/mkt/event58539_manage.asp?returndate=2015-01-25">2015-01-25</a></td>
		<td bgcolor="<%=chkiif(Date()="2015-01-26","FFFF09","")%>"><a href="http://webadmin.10x10.co.kr/admin/datamart/mkt/event58539_manage.asp?returndate=2015-01-26">2015-01-26</a></td>
		<td bgcolor="<%=chkiif(Date()="2015-01-27","FFFF09","")%>"><a href="http://webadmin.10x10.co.kr/admin/datamart/mkt/event58539_manage.asp?returndate=2015-01-27">2015-01-27</a></td>
		<td bgcolor="<%=chkiif(Date()="2015-01-28","FFFF09","")%>"><a href="http://webadmin.10x10.co.kr/admin/datamart/mkt/event58539_manage.asp?returndate=2015-01-28">2015-01-28</a></td>
		<td bgcolor="<%=chkiif(Date()="2015-01-29","FFFF09","")%>"><a href="http://webadmin.10x10.co.kr/admin/datamart/mkt/event58539_manage.asp?returndate=2015-01-29">2015-01-29</a></td>
		<td bgcolor="<%=chkiif(Date()="2015-01-30","FFFF09","")%>"><a href="http://webadmin.10x10.co.kr/admin/datamart/mkt/event58539_manage.asp?returndate=2015-01-30">2015-01-30</a></td>
	</tr>
	
</table>


<table style="margin:0 auto;text-align:center;margin-top:30px;">
	<tr>
		<td>현재 보고 있는 날짜 : <%=returndate%>일 &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<strong>어벤져박스 시간대 별 통계</strong></td>
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
		<th><strong>A-Main 클릭수 </strong></th>
		<th><strong>kakao 클리수</strong></th>
		<th><strong>기획전 1번</strong></th>
		<th><strong>기획전 2번</strong></th>
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
						" 		,isnull(l.[banner1],0) as banner1  "&_
						" 		,isnull(l.[banner2],0) as banner2  "&_
						" 	FROM [Hours] h  "&_
						" 	LEFT OUTER JOIN ( "&_
						" 		select   "&_
						" 			datepart(hh,regdate) as hhtime   "&_
						" 			, count(case when chkid = 'mo_main' then chkid end) as mo_main  "&_
						" 			, count(case when chkid = 'app_main' then chkid end) as app_main  "&_
						" 			, count(case when chkid = 'kakao' then chkid end) as kakao  "&_
						" 			, count(case when chkid = 'banner1' then chkid end) as banner1  "&_
						" 			, count(case when chkid = 'banner2' then chkid end) as banner2  "&_
						" 		from db_datamart.dbo.tbl_event_click_log   "&_
						" 		where eventid='58539'   "&_
						" 		and convert(varchar(10),regdate,120) between '"& returndate &"' and '"& returndate &"'  "&_
						" 		group by datepart(hh,regdate) "&_
						" 	) as l "&_
						" ON h.[Hour]-1 = l.[hhtime] " &_
						" order by h.[Hour] asc "
				
				'Response.write sqlStr

				db3_rsget.Open sqlStr,db3_dbget
				Dim mo_main , app_main , kakao , banner1 , banner2
				mo_main = 0
				app_main = 0
				kakao = 0
				banner1 = 0
				banner2 = 0
				if Not(db3_rsget.EOF or db3_rsget.BOF) Then
					Do Until db3_rsget.eof
			%>
			<tr style="text-align:center;">
				<td width="15%" bgcolor="<%=chkiif(db3_rsget("mo_main")>0,"skyblue","")%>"><%=db3_rsget("mo_main")%></td>
				<td width="15%" bgcolor="<%=chkiif(db3_rsget("app_main")>0,"skyblue","")%>"><%=db3_rsget("app_main")%></td>
				<td width="15%" bgcolor="<%=chkiif(db3_rsget("kakao")>0,"skyblue","")%>"><%=db3_rsget("kakao")%></td>
				<td width="15%" bgcolor="<%=chkiif(db3_rsget("banner1")>0,"skyblue","")%>"><%=db3_rsget("banner1")%></td>
				<td width="15%" bgcolor="<%=chkiif(db3_rsget("banner2")>0,"skyblue","")%>"><%=db3_rsget("banner2")%></td>
			</tr>
			<%
					mo_main = mo_main + db3_rsget("mo_main")
					app_main = app_main + db3_rsget("app_main")
					kakao = kakao + db3_rsget("kakao")
					banner1 = banner1 + db3_rsget("banner1")
					banner2 = banner2 + db3_rsget("banner2")
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
	<td width="12.5%"><%=banner1%></td>
	<td width="12.5%"><%=banner2%></td>
	<td width="12.5%"><%=usercnt%></td>
	<td width="12.5%"><%=ordercnt%></td>
</tr>
</table>

<% If session("ssBctId")="gawisonten10" or session("ssBctId") = "motions" Or session("ssBctId") = "greenteenz" Then %>
<!-- 특정 시간대 구매 데이터 11:30 ~11:40 구매자 -->
<table style="margin:0 auto;text-align:center;margin-top:30px;">
	<tr>
		<td>현재 보고 있는 날짜 : <%=returndate%>일 &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<strong>어벤져박스 특정 시간대 구매자</strong></td>
	</tr>
</table>
<table class="table" style="width:90%;">
	<colgroup>
		<col width="10%" />
		<col width="*" />
		<col width="*" />
	</colgroup>
	<tr align="center" bgcolor="#E6E6E6">
		<th><strong>주문번호</strong></th>
		<th><strong>회원아이디</strong></th>
		<th><strong>주문시간대</strong></th>
	</tr>
	<%
		sqlStr = " select m.orderserial , m.userid , m.regdate from db_order.dbo.tbl_order_master as m "&_
				" inner join db_order.dbo.tbl_order_detail as d "&_
				" on m.orderserial=d.orderserial "&_
				" where m.jumundiv<>'9' and m.ipkumdiv > 3 and m.cancelyn = 'N'  "&_
				" and d.cancelyn<>'Y' and d.itemid<>'0' and d.itemid = '"& itemid &"' "&_
				" and m.regdate between '"&returndate&" 11:30:00' and '"&returndate&" 11:40:00' "

		db3_rsget.Open sqlStr,dbget
		if Not(db3_rsget.EOF or db3_rsget.BOF) Then
			Do Until db3_rsget.eof
	%>
	<tr bgcolor="#FFFFFF" align="center">
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
<!-- 특정 시간대 구매 데이터 11:30 ~11:40 구매자 -->
<% End If %>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->
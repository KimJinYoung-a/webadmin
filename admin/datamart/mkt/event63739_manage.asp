<%@ language=vbscript %>
<% option Explicit %>
<% Response.CharSet = "euc-kr" %>
<%
'####################################################
' Description : 아이커피 1차
' History : 2015.06.18 한용민 생성
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<%
dim eCode, subscriptcount, userid, evtTotalCnt, sqlStr
Dim returndate  : returndate = 	request("returndate")

If returndate = "" Then returndate = Date()
'returndate = "2015-06-19"
	
IF application("Svr_Info") = "Dev" THEN
	eCode 		= "63794"
Else
	eCode 		= "63739"
End If

If session("ssBctId")="winnie" Or session("ssBctId")="gawisonten10" Or session("ssBctId") = "edojun" Or session("ssBctId") = "tozzinet" Or session("ssBctId") = "thensi7" Or session("ssBctId") = "bborami" Or session("ssBctId")="stella0117" Or session("ssBctId")="jinyeonmi" Or session("ssBctId")="kyungae13" Then
Else
	response.write "<script type='text/javascript'>alert('관계자만 볼 수 있는 페이지 입니다.'); window.close();</script>"
	dbget.close() : Response.End
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
		
		winnumberfrm.action="/admin/datamart/mkt/event63739_manage_process.asp"
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
		<strong><font size=5>아이러브XX 시리즈 관리자<br>현재 보고 있는 날짜 : <%=returndate%>일</font></strong>
		<br><span style="color:red;">부하가 가는 페이지입니다. 막 누르지 마세요.</span>
	</td>
</tr>
</table>

<table style="margin:0 auto;text-align:center;margin-top:50px;">
<tr>
	<td>
		<span style="color:red;">
			1부터 - 1000 사이 비율로 확률을 입력해 주시면 됩니다.
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
<tr bgcolor="#FFFFFF" align="center">
	<td>
<%
dim limitcoffecnt
limitcoffecnt=0

sqlStr = "select count(*) as cnt"
sqlStr = sqlStr & " from [db_event].[dbo].[tbl_event_subscript] sc"
sqlstr = sqlstr & " where sc.evt_code="& eCode &""
sqlstr = sqlstr & " and isnull(sc.sub_opt1,'') = '"& returndate &"'"
sqlstr = sqlstr & " and isnull(sc.sub_opt2,'') = '1'"

'response.write sqlStr & "<br>"
rsget.Open sqlStr,dbget
if Not(rsget.EOF or rsget.BOF) Then
	limitcoffecnt=rsget("cnt")
else
	limitcoffecnt=0
End If
rsget.close
%>
1차:스타벅스 / 2차:비비큐 / 3차:떡볶이 발급수 : <%= limitcoffecnt %>

<%
dim limitmachinecnt
limitmachinecnt=0

sqlStr = "select count(*) as cnt"
sqlStr = sqlStr & " from [db_event].[dbo].[tbl_event_subscript] sc"
sqlstr = sqlstr & " where sc.evt_code="& eCode &""
sqlstr = sqlstr & " and isnull(sc.sub_opt1,'') = '"& returndate &"'"
sqlstr = sqlstr & " and isnull(sc.sub_opt2,'') = '2'"

'response.write sqlStr & "<br>"
rsget.Open sqlStr,dbget
if Not(rsget.EOF or rsget.BOF) Then
	limitmachinecnt=rsget("cnt")
else
	limitmachinecnt=0
End If
rsget.close
%>
<br>1차:커피메이커 / 2차:타올 / 3차:반8 발급수 : <%= limitmachinecnt %>

<%
dim coupon1, couponidx_1
coupon1=0
if returndate>="2015-06-26" then
	couponidx_1 = 10418
elseif returndate>="2015-06-23" then
	couponidx_1 = 10413
else
	couponidx_1 = 10404
end if

sqlStr = "select count(*) as cnt"
sqlStr = sqlStr & " from [db_event].[dbo].[tbl_event_subscript] sc"
sqlstr = sqlstr & " where sc.evt_code="& eCode &""
sqlstr = sqlstr & " and isnull(sc.sub_opt1,'') = '"& returndate &"'"
sqlstr = sqlstr & " and isnull(sc.sub_opt2,'') = '"& couponidx_1 &"'"

'response.write sqlStr & "<br>"
rsget.Open sqlStr,dbget
if Not(rsget.EOF or rsget.BOF) Then
	coupon1=rsget("cnt")
else
	coupon1=0
End If
rsget.close
%>
<br><br>쿠폰발급수(<%= couponidx_1 %>) : <%= coupon1 %>

<%
dim coupon2, couponidx_2
coupon2=0
if returndate>="2015-06-26" then
	couponidx_2 = 10419
elseif returndate>="2015-06-23" then
	couponidx_2 = 10414
else
	couponidx_2 = 10405
end if

sqlStr = "select count(*) as cnt"
sqlStr = sqlStr & " from [db_event].[dbo].[tbl_event_subscript] sc"
sqlstr = sqlstr & " where sc.evt_code="& eCode &""
sqlstr = sqlstr & " and isnull(sc.sub_opt1,'') = '"& returndate &"'"
sqlstr = sqlstr & " and isnull(sc.sub_opt2,'') = '"& couponidx_2 &"'"

'response.write sqlStr & "<br>"
rsget.Open sqlStr,dbget
if Not(rsget.EOF or rsget.BOF) Then
	coupon2=rsget("cnt")
else
	coupon2=0
End If
rsget.close
%>
<br>쿠폰발급수(<%= couponidx_2 %>) : <%= coupon2 %>

<%
dim coupon3, couponidx_3
coupon3=0
if returndate>="2015-06-26" then
	couponidx_3 = 10420
elseif returndate>="2015-06-23" then
	couponidx_3 = 10415
else
	couponidx_3 = 10406
end if

sqlStr = "select count(*) as cnt"
sqlStr = sqlStr & " from [db_event].[dbo].[tbl_event_subscript] sc"
sqlstr = sqlstr & " where sc.evt_code="& eCode &""
sqlstr = sqlstr & " and isnull(sc.sub_opt1,'') = '"& returndate &"'"
sqlstr = sqlstr & " and isnull(sc.sub_opt2,'') = '"& couponidx_3 &"'"

'response.write sqlStr & "<br>"
rsget.Open sqlStr,dbget
if Not(rsget.EOF or rsget.BOF) Then
	coupon3=rsget("cnt")
else
	coupon3=0
End If
rsget.close
%>
<br>쿠폰발급수(<%= couponidx_3 %>) : <%= coupon3 %>

<%
dim coupon4, couponidx_4
coupon4=0
if returndate>="2015-06-26" then
	couponidx_4 = 10421
elseif returndate>="2015-06-23" then
	couponidx_4 = 10416
else
	couponidx_4 = 10407
end if

sqlStr = "select count(*) as cnt"
sqlStr = sqlStr & " from [db_event].[dbo].[tbl_event_subscript] sc"
sqlstr = sqlstr & " where sc.evt_code="& eCode &""
sqlstr = sqlstr & " and isnull(sc.sub_opt1,'') = '"& returndate &"'"
sqlstr = sqlstr & " and isnull(sc.sub_opt2,'') = '"& couponidx_4 &"'"

'response.write sqlStr & "<br>"
rsget.Open sqlStr,dbget
if Not(rsget.EOF or rsget.BOF) Then
	coupon4=rsget("cnt")
else
	coupon4=0
End If
rsget.close
%>
<br>쿠폰발급수(<%= couponidx_4 %>) : <%= coupon4 %>
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
	<td><a href="http://webadmin.10x10.co.kr/admin/datamart/mkt/event63739_manage.asp?returndate=2015-06-19">2015-06-19</a></td>
	<td><a href="http://webadmin.10x10.co.kr/admin/datamart/mkt/event63739_manage.asp?returndate=2015-06-20">2015-06-20</a></td>
	<td><a href="http://webadmin.10x10.co.kr/admin/datamart/mkt/event63739_manage.asp?returndate=2015-06-21">2015-06-21</a></td>
	<td><a href="http://webadmin.10x10.co.kr/admin/datamart/mkt/event63739_manage.asp?returndate=2015-06-22">2015-06-22</a></td>
</tr>																				            
<tr bgcolor="#FFFFFF" align="center">
	<td><a href="http://webadmin.10x10.co.kr/admin/datamart/mkt/event63739_manage.asp?returndate=2015-06-23">2015-06-23</a></td>
	<td><a href="http://webadmin.10x10.co.kr/admin/datamart/mkt/event63739_manage.asp?returndate=2015-06-24">2015-06-24</a></td>
	<td><a href="http://webadmin.10x10.co.kr/admin/datamart/mkt/event63739_manage.asp?returndate=2015-06-25">2015-06-25</a></td>
	<td></td>
</tr>	
<tr bgcolor="#FFFFFF" align="center">
	<td><a href="http://webadmin.10x10.co.kr/admin/datamart/mkt/event63739_manage.asp?returndate=2015-06-26">2015-06-26</a></td>
	<td><a href="http://webadmin.10x10.co.kr/admin/datamart/mkt/event63739_manage.asp?returndate=2015-06-27">2015-06-27</a></td>
	<td><a href="http://webadmin.10x10.co.kr/admin/datamart/mkt/event63739_manage.asp?returndate=2015-06-28">2015-06-28</a></td>
	<td></td>
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
</colgroup>
<tr align="center" bgcolor="#E6E6E6">
	<th><strong>시간</strong></th>
	<th><strong>M-앱다운받기 클릭수</strong></th>
	<th><strong>A-전면배너 클릭수</strong></th>
	<th><strong>카톡 클릭수</strong></th>
	<th><strong>회원가입수</strong></th>
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
					" 			, count(case when chkid = 'KAKAO' then chkid end) as kakao  "&_
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
</tr>
</table>
<table class="table" style="width:90%; text-align:center;" border="0" cellspacing="0" cellpadding="0">
<tr>
	<td width="12.5%">합계</td>
	<td width="12.5%"><%=mo_main%></td>
	<td width="12.5%"><%=app_main%></td>
	<td width="12.5%"><%=kakao%></td>
	<td width="12.5%"><%=usercnt%></td>
</tr>
</table>

<% If session("ssBctId")="winnie" Or session("ssBctId")="gawisonten10" Or session("ssBctId") = "edojun" Or session("ssBctId") = "tozzinet" Or session("ssBctId") = "thensi7" Or session("ssBctId") = "bborami" Or session("ssBctId")="stella0117" Or session("ssBctId")="jinyeonmi" Or session("ssBctId")="kyungae13" Then %>
	<table style="margin:0 auto;text-align:center;margin-top:30px;">
		<tr>
			<td>1차:스타벅스 / 2차:비비큐 / 3차:떡볶이 당첨자 선착순 1000명 ※ 약 30분 정도의 오차가 있습니다.</td>
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
		<th><strong>순위</strong></th>
		<th><strong>회원아이디</strong></th>
		<th><strong>핸드폰번호</strong></th>
	</tr>
	<%
		dim tmpval
		tmpval = 0
		sqlStr = " select top 1100"&_
				" s.userid, n.usercell"&_
				" from [db_event].[dbo].[tbl_event_subscript] s"&_
				" left join db_user.dbo.tbl_user_n n"&_
				" 	on s.userid=n.userid "&_
				" where s.evt_code="& eCode &""&_
				" and isnull(s.sub_opt1,'') = '"& returndate &"'"&_
				" and isnull(s.sub_opt2,'') = '1'"&_
				" order by s.sub_idx asc"

		'response.write sqlStr & "<br>"
		db3_rsget.Open sqlStr,dbget
		if Not(db3_rsget.EOF or db3_rsget.BOF) Then
			Do Until db3_rsget.eof
			tmpval = tmpval + 1
	%>
	<tr bgcolor="#FFFFFF" align="center">
		<td><%= tmpval %></td>
		<td><%=db3_rsget("userid")%></td>
		<td><%=db3_rsget("usercell")%></td>
	</tr>
	<%
			db3_rsget.movenext
			Loop
		End If
		db3_rsget.close
	%>
	</table>
	
	<table style="margin:0 auto;text-align:center;margin-top:30px;">
		<tr>
			<td>1차:커피메이커 / 2차:계란후라이타올 / 3차:반8 당첨자 선착순 10명 ※ 약 30분 정도의 오차가 있습니다.</td>
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
		<th><strong>순위</strong></th>
		<th><strong>회원아이디</strong></th>
		<th><strong>핸드폰번호</strong></th>
	</tr>
	<%
		tmpval = 0
		sqlStr = " select top 20"&_
				" s.userid, n.usercell"&_
				" from [db_event].[dbo].[tbl_event_subscript] s"&_
				" left join db_user.dbo.tbl_user_n n"&_
				" 	on s.userid=n.userid "&_
				" where s.evt_code="& eCode &""&_
				" and isnull(s.sub_opt1,'') = '"& returndate &"'"&_
				" and isnull(s.sub_opt2,'') = '2'"&_
				" order by s.sub_idx asc"

		'response.write sqlStr & "<br>"
		db3_rsget.Open sqlStr,dbget
		if Not(db3_rsget.EOF or db3_rsget.BOF) Then
			Do Until db3_rsget.eof
			tmpval = tmpval + 1
	%>
	<tr bgcolor="#FFFFFF" align="center">
		<td><%= tmpval %></td>
		<td><%=db3_rsget("userid")%></td>
		<td><%=db3_rsget("usercell")%></td>
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

<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->
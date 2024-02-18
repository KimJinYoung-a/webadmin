<%@ language=vbscript %>
<% option Explicit %>
<% Response.CharSet = "euc-kr" %>
<%
'####################################################
' Description : 돌아온 크리스박스의 기적!
' History : 2015.12.07 유태욱 생성
'####################################################
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
	eCode 		= "65978"
	itemid		= getdateitem(returndate)
Else
	eCode 		= "67929"
	itemid		= getdateitem(returndate)
End If

If session("ssBctId")="winnie" Or session("ssBctId")="gawisonten10" Or session("ssBctId") ="greenteenz" Or session("ssBctId") = "edojun" Or session("ssBctId") = "tozzinet" Or session("ssBctId") = "thensi7" Then

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
		
		winnumberfrm.action="/admin/datamart/mkt/event67929_manage_process.asp"
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
		<strong><font size=5>크리스박스 관리자<br>현재 보고 있는 날짜 : <%=returndate%>일</font></strong>
		<br><br><b><span style="color:red;">부하가 가는 페이지입니다. 막 누르지 마세요.</span></b>
	</td>
</tr>
</table>

<table style="margin:0 auto;text-align:center;margin-top:50px;">
<tr>
	<td>
		<span style="color:red;">
			1부터 - 10000 사이 비율로 입력해 주시면 됩니다.
			<br>0.01% -> 2 입력
			<br>0.1% -> 11 입력
			<br>1% -> 101 입력
			<br>10% -> 1001 입력
			<br>50% -> 5001 입력
			<br>90% -> 9001 입력
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

<%
	dim cnt, allcnt
	''총 응모자수

	sqlstr = "select count(userid) as cnt, SUM(CONVERT(BIGINT, sub_opt1)) as allcnt " &_
			"  from db_event.dbo.tbl_event_subscript where evt_code = '" & eCode & "' and regdate between '"&returndate&" 00:00:00' and '"&returndate&" 23:59:59' "
			'response.write sqlstr
	rsget.Open sqlStr,dbget,1
		cnt = rsget(0)
		allcnt = rsget(1)
	rsget.Close
%>

<tr bgcolor="#FFFFFF" align="center">
	<td><a href="http://webadmin.10x10.co.kr/admin/datamart/mkt/event67929_manage.asp?returndate=2015-12-09">2015-12-09</a></td>
	<td><a href="http://webadmin.10x10.co.kr/admin/datamart/mkt/event67929_manage.asp?returndate=2015-12-10">2015-12-10</a></td>
	<td><a href="http://webadmin.10x10.co.kr/admin/datamart/mkt/event67929_manage.asp?returndate=2015-12-11">2015-12-11</a></td>
	<td><a href="http://webadmin.10x10.co.kr/admin/datamart/mkt/event67929_manage.asp?returndate=2015-12-16">2015-12-16</a></td>
	<td><a href="http://webadmin.10x10.co.kr/admin/datamart/mkt/event67929_manage.asp?returndate=2015-12-17">2015-12-17</a></td>
	<td><a href="http://webadmin.10x10.co.kr/admin/datamart/mkt/event67929_manage.asp?returndate=2015-12-18">2015-12-18</a></td>
</tr>																				            
</table>

<table style="margin:0 auto;text-align:center;margin-top:30px;">
<tr>
	<td><font color="RED">응모자수 : <%= cnt %></font></td>
</tr>
<tr>
	<td><font color="RED">응모건수 : <%= allcnt %></font></td>
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
	<th><strong>앱다운 배너 클릭수</strong></th>
	<th><strong>앱 전면배너 클릭수</strong></th>
	<th><strong>kakao 클릭수</strong></th>
	<th><strong>facebook 클릭수</strong></th>
	<th><strong>line 클릭수</strong></th>
	<th><strong>회원가입수</strong></th>
	<th><strong>상품판매수</strong></th>
</tr>
</table>
<table class="table" style="width:90%;" border="0" cellspacing="0" cellpadding="0">
<tr>
	<td width="7%">
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
	<td width="37.5%"> 
		<table width="100%">
		<%
			sqlStr = " ; WITH [Hours] ([Hour]) AS "&_
					" (  "&_
					" SELECT TOP 24 ROW_NUMBER() OVER (ORDER BY [object_id]) AS [Hour]  "&_
					" FROM sys.objects  "&_
					" ORDER BY [object_id]  "&_
					" )  "&_
					" 	SELECT h.[Hour]-1 as Hour  "&_
					" 		,isnull(l.[appdncnt],0) as appdncnt  "&_
					" 		,isnull(l.[app_main],0) as app_main  "&_
					" 		,isnull(l.[kakao],0) as kakao  "&_
					" 		,isnull(l.[facebook],0) as facebook  "&_
					" 		,isnull(l.[linemsg],0) as linemsg  "&_
					" 	FROM [Hours] h  "&_
					" 	LEFT OUTER JOIN ( "&_
					" 		select   "&_
					" 			datepart(hh,regdate) as hhtime   "&_
					" 			, count(case when chkid = 'appdncnt' then chkid end) as appdncnt  "&_
					" 			, count(case when chkid = 'app_main' then chkid end) as app_main  "&_
					" 			, count(case when chkid = 'kk' then chkid end) as kakao  "&_
					" 			, count(case when chkid = 'fb' then chkid end) as facebook  "&_
					" 			, count(case when chkid = 'ln' then chkid end) as linemsg  "&_
					" 		from db_datamart.dbo.tbl_event_click_log   "&_
					" 		where eventid='"& eCode &"'   "&_
					" 		and convert(varchar(10),regdate,120) between '"& returndate &"' and '"& returndate &"'  "&_
					" 		group by datepart(hh,regdate) "&_
					" 	) as l "&_
					" ON h.[Hour]-1 = l.[hhtime] " &_
					" order by h.[Hour] asc "
			
			'Response.write sqlStr

			db3_rsget.Open sqlStr,db3_dbget
			Dim appdncnt , app_main , kakao, facebook, linemsg
			appdncnt = 0
			app_main = 0
			kakao = 0
			facebook = 0
			linemsg = 0

			if Not(db3_rsget.EOF or db3_rsget.BOF) Then
				Do Until db3_rsget.eof
		%>
		<tr style="text-align:center;">
			<td width="15%" bgcolor="<%=chkiif(db3_rsget("appdncnt")>0,"skyblue","")%>"><%=db3_rsget("appdncnt")%></td>
			<td width="15%" bgcolor="<%=chkiif(db3_rsget("app_main")>0,"skyblue","")%>"><%=db3_rsget("app_main")%></td>
			<td width="15%" bgcolor="<%=chkiif(db3_rsget("kakao")>0,"skyblue","")%>"><%=db3_rsget("kakao")%></td>
			<td width="15%" bgcolor="<%=chkiif(db3_rsget("facebook")>0,"skyblue","")%>"><%=db3_rsget("facebook")%></td>
			<td width="15%" bgcolor="<%=chkiif(db3_rsget("linemsg")>0,"skyblue","")%>"><%=db3_rsget("linemsg")%></td>
		</tr>
		<%
				appdncnt = appdncnt + db3_rsget("appdncnt")
				app_main = app_main + db3_rsget("app_main")
				kakao = kakao + db3_rsget("kakao")
				facebook = facebook + db3_rsget("facebook")
				linemsg = linemsg + db3_rsget("linemsg")

				db3_rsget.movenext
				Loop
			End If
			db3_rsget.close
		%>
		</table>
	</td>
	<td width="7%">
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
	<td width="7%">
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
	<td width="12.5%"><%=appdncnt%></td>
	<td width="12.5%"><%=app_main%></td>
	<td width="12.5%"><%=kakao%></td>
	<td width="12.5%"><%=facebook%></td>
	<td width="12.5%"><%=linemsg%></td>
	<td width="12.5%"><%=usercnt%></td>
	<td width="12.5%"><%=ordercnt%></td>
</tr>
</table>
 
<% If session("ssBctId")="gawisonten10" or session("ssBctId") = "motions" Or session("ssBctId") = "greenteenz" Or session("ssBctId") = "thensi7" Or session("ssBctId") = "tozzinet" Then %>
	<table style="margin:0 auto;text-align:center;margin-top:30px;">
		<tr>
			<td>얼리버드 예약자</td>
		</tr>
		<%
			dim birdcnt
			''얼리버드 응모자수
		
			sqlstr = "select COUNT(DISTINCT userid) " &_
					"  from db_log.[dbo].[tbl_caution_event_log] where evt_code = '" & eCode & "' and value1='al_com' and regdate between '"&returndate&" 00:00:00' and '"&returndate&" 23:59:59' "
					'response.write sqlstr
			rsget.Open sqlStr,dbget,1
				birdcnt = rsget(0)
			rsget.Close
		%>
		<tr>
			<td>
				<%= birdcnt %>
			</td>
		</tr>
	</table>
<!--	
	<table class="table" style="width:90%;">
	<colgroup>
		<col width="5%" />
		<col width="10%" />
		<col width="*" />
		<col width="*" />
	</colgroup>
	<tr align="center" bgcolor="#E6E6E6">
		<th><strong>순번</strong></th>
		<th><strong>회원아이디</strong></th>
		<th><strong>예약시간</strong></th>
	</tr>
-->
	<%
'		dim tmpval
'		tmpval = 0
		''기존
'		sqlStr = " select top 1000 m.orderserial , m.userid , m.regdate from db_order.dbo.tbl_order_master as m "&_
'				" inner join db_order.dbo.tbl_order_detail as d "&_
'				" on m.orderserial=d.orderserial "&_
'				" where m.jumundiv<>'9' and m.ipkumdiv > 3 and m.cancelyn = 'N'  "&_
'				" and d.cancelyn<>'Y' and d.itemid<>'0' and d.itemid = '"& itemid &"' "&_
'				" and m.regdate between '"&returndate&" 10:00:00' and '"&returndate&" 23:59:59' "&_
'				" order by m.orderserial asc"

		''신규
'		sqlStr = " select userid, regdate from db_log.[dbo].[tbl_caution_event_log] "&_
'				" where evt_code='"&eCode&"' "&_
'				" and value1='al_com' and regdate between '"&returndate&" 00:00:00' and '"&returndate&" 23:59:59' "&_
'				" order by idx asc"
'
'		db3_rsget.Open sqlStr,dbget
'		if Not(db3_rsget.EOF or db3_rsget.BOF) Then
'			Do Until db3_rsget.eof
'			tmpval = tmpval + 1
	%>
<!--
	<tr bgcolor="#FFFFFF" align="center">
		<td> tmpval </td>
		<td>db3_rsget("userid")</td>
		<td>db3_rsget("regdate")</td>
	</tr>
-->
	<%
'			db3_rsget.movenext
'			Loop
'		End If
'		db3_rsget.close
	%>
	</table>

<!--
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
-->
	<%
'		dim tmpval
'		tmpval = 0
'		sqlStr = " select top 1000 m.orderserial , m.userid , m.regdate from db_order.dbo.tbl_order_master as m "&_
'				" inner join db_order.dbo.tbl_order_detail as d "&_
'				" on m.orderserial=d.orderserial "&_
'				" where m.jumundiv<>'9' and m.ipkumdiv > 3 and m.cancelyn = 'N'  "&_
'				" and d.cancelyn<>'Y' and d.itemid<>'0' and d.itemid = '"& itemid &"' "&_
'				" and m.regdate between '"&returndate&" 10:00:00' and '"&returndate&" 23:59:59' "&_
'				" order by m.orderserial asc"
'
'		db3_rsget.Open sqlStr,dbget
'		if Not(db3_rsget.EOF or db3_rsget.BOF) Then
'			Do Until db3_rsget.eof
'			tmpval = tmpval + 1
	%>
<!--
	<tr bgcolor="#FFFFFF" align="center">

	</tr>
-->
	<%
'			db3_rsget.movenext
'			Loop
'		End If
'		db3_rsget.close
	%>
	<!--</table>-->

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
		if dateval="2015-12-09" then
			tmpdateitem=1212183
		elseif dateval="2015-12-10" then
			tmpdateitem=1212183
		elseif dateval="2015-12-11" then
			tmpdateitem=1212183

		elseif dateval="2015-12-16" then
			tmpdateitem=1212183
		elseif dateval="2015-12-17" then
			tmpdateitem=1212183
		elseif dateval="2015-12-18" then
			tmpdateitem=1212183
		else
			tmpdateitem=""
		end if
	Else
		if dateval="2015-12-09" then
			tmpdateitem=1404138
		elseif dateval="2015-12-10" then
			tmpdateitem=1404138
		elseif dateval="2015-12-11" then
			tmpdateitem=1404138

		elseif dateval="2015-12-16" then
			tmpdateitem=1404911
		elseif dateval="2015-12-17" then
			tmpdateitem=1404911
		elseif dateval="2015-12-18" then
			tmpdateitem=1404911
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
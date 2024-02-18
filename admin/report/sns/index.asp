<%@ language=vbscript %>
<% option explicit %>
<% response.Buffer=true %>
<%
'###########################################################
' Description :  SNS 회원가입 통계
' History : 2014.07.07 유태욱
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/partner/lib/adminHead.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/partner/lib/function/incPageFunction.asp" -->
<!-- #Include virtual = "/lib/classes/report/snsdataCls.asp" -->
<div align="center" id="loadingBar" style="width:100%; height100%;">
<img src="https://webadmin.10x10.co.kr/images/loading2.gif">
</div>
<%
Response.Flush
	dim joindate, joinyy, joinmm, joinww, nowww, datename, forcnt, fordate
	joindate	= requestCheckVar(request("joindate"),2)	'날짜구분(년도별,월별,주차별,일별)
	joinyy		= requestCheckVar(request("joinyy"),4)	'년도선택
	joinmm		= requestCheckVar(request("joinmm"),2)	'월선택
	joinww		= requestCheckVar(request("joinww"),2)	'분기선택

	if joindate="" or isNull(joindate) then
		joindate = "dd"
	end if
	
	if joinyy="" or isNull(joinyy) then
		joinyy = Year(Date())
	end if

	if joinmm="" or isNull(joinmm) then
		joinmm = Month(Date())
	end if

	if joinww="" or isNull(joinww) then
		nowww = DATEPART("ww", Date())
		if nowww < 14 then
			joinww = 1
		elseif nowww >= 14 and nowww <= 26 then
			joinww = 2
		elseif nowww >= 27 and nowww <= 39 then
			joinww = 3
		else
			joinww = 4
		end if
	end if

	if joindate = "dd" then
		datename = "일"
		forcnt = 30
		fordate = 1
	elseif joindate = "ww" then
		datename = "주차"
		forcnt = 12
		if joinww = 1 then
			fordate = 1
		elseif joinww = 2 then
			fordate = 14
		elseif joinww = 3 then
			fordate = 27
		elseif joinww = 4 then
			fordate = 40
		end if
		
	elseif joindate = "mm" then
		datename = "월"
		forcnt = 11
		fordate = 1
	elseif joindate = "yy" then
		datename = "년"
		forcnt = 3
		fordate = joinyy-forcnt
	end if

	'// 내용 접수
	dim oSnsJoinList, i, alldata
	Set oSnsJoinList = New CSnsContents
		oSnsJoinList.FRectjoindate	= joindate		'년도,월별,주차,일별
		oSnsJoinList.FRectjoinyy	= joinyy		'년도선택
		oSnsJoinList.FRectjoinmm	= joinmm		'월선택
		oSnsJoinList.FRectjoinww	= joinww		'분기선택(주차별)
		alldata = oSnsJoinList.GetSnsjoinList()
	set oSnsJoinList = Nothing

%>

<style>
input[type=radio] {
		display:none;
	}

input[type=radio] + label {
	display:inline-block;
	margin:-3px;
	padding: 8px 12px;
	margin-bottom: 0;
	font-size: 14px;
	line-height: 20px;
	color: #333;
	text-align: center;
	text-shadow: 0 1px 1px rgba(255,255,255,0.75);
	vertical-align: middle;
	cursor: pointer;
	background-color: #f5f5f5;
	background-image: -moz-linear-gradient(top,#fff,#e6e6e6);
	background-image: -webkit-gradient(linear,0 0,0 100%,from(#fff),to(#e6e6e6));
	background-image: -webkit-linear-gradient(top,#fff,#e6e6e6);
	background-image: -o-linear-gradient(top,#fff,#e6e6e6);
	background-image: linear-gradient(to bottom,#fff,#e6e6e6);
	background-repeat: repeat-x;
	border: 1px solid #ccc;
	border-color: #e6e6e6 #e6e6e6 #bfbfbf;
	border-color: rgba(0,0,0,0.1) rgba(0,0,0,0.1) rgba(0,0,0,0.25);
	border-bottom-color: #b3b3b3;
	filter: progid:DXImageTransform.Microsoft.gradient(startColorstr='#ffffffff',endColorstr='#ffe6e6e6',GradientType=0);
	filter: progid:DXImageTransform.Microsoft.gradient(enabled=false);
	-webkit-box-shadow: inset 0 1px 0 rgba(255,255,255,0.2),0 1px 2px rgba(0,0,0,0.05);
	-moz-box-shadow: inset 0 1px 0 rgba(255,255,255,0.2),0 1px 2px rgba(0,0,0,0.05);
	box-shadow: inset 0 1px 0 rgba(255,255,255,0.2),0 1px 2px rgba(0,0,0,0.05);
}

input[type=radio]:checked + label {
	background-image: none;
	outline: 0;
	-webkit-box-shadow: inset 0 2px 4px rgba(0,0,0,0.15),0 1px 2px rgba(0,0,0,0.05);
	-moz-box-shadow: inset 0 2px 4px rgba(0,0,0,0.15),0 1px 2px rgba(0,0,0,0.05);
	box-shadow: inset 0 2px 4px rgba(0,0,0,0.15),0 1px 2px rgba(0,0,0,0.05);
	background-color:#e0e0e0;
}

.selectBox {
	height: 38px;
	margin:-3px;
	padding: 8px 12px;
	margin-bottom: 0;
	font-size: 14px;
	line-height: 20px;
	color: #333;
	text-align: center;
	text-shadow: 0 1px 1px rgba(255,255,255,0.75);
	vertical-align: middle;
	cursor: pointer;
	background-color: #e0e0e0;
	background-image: none;
	outline: 0;
	border: 1px solid #ccc;
	border-color: #e6e6e6 #e6e6e6 #bfbfbf;
	border-color: rgba(0,0,0,0.1) rgba(0,0,0,0.1) rgba(0,0,0,0.25);
	border-bottom-color: #b3b3b3;
	filter: progid:DXImageTransform.Microsoft.gradient(startColorstr='#ffffffff',endColorstr='#ffe6e6e6',GradientType=0);
	filter: progid:DXImageTransform.Microsoft.gradient(enabled=false);
	-webkit-box-shadow: inset 0 2px 4px rgba(0,0,0,0.15),0 1px 2px rgba(0,0,0,0.05);
	-moz-box-shadow: inset 0 2px 4px rgba(0,0,0,0.15),0 1px 2px rgba(0,0,0,0.05);
	box-shadow: inset 0 2px 4px rgba(0,0,0,0.15),0 1px 2px rgba(0,0,0,0.05);
	-webkit-appearance:none; /* for chrome */
	-moz-appearance:none; /*for firefox*/
	appearance:none;
}
</style>

<script language="javascript">
function searchFrm(){
	frm.submit();
}

function fndateselect(sel){
	if(sel=='mm'){
		$('#month').hide();
		$('#weeklist').hide();
	}else if(sel=='ww'){
		$('#month').hide();
		$('#weeklist').show();
	}else if(sel=='dd'){
		$('#month').show();
		$('#weeklist').hide();
	}else if(sel=='yy'){
		$('#month').hide();
		$('#weeklist').hide();
	}
}

function colorCH(cell,color){
	if(color=="on"){
		$(cell).addClass("bgBl1");
	}else if(color=="off"){
		$(cell).removeClass("bgBl1");
	}
}
</script>

</head>
<body>
<div class="wrap"><br><br>
	<!-- search -->
	<form name="frm" method="get" action="" action="index.asp">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<div class="searchWrap">
		<ul id="daylist"><!--  style="display:none;" -->
			<input type="radio" id="radioyy" onclick="fndateselect('yy');" name="joindate" value="yy" <% if joindate="yy" then %>checked<% end if %>>
			<label for="radioyy">년도별</label>

			<input type="radio" id="radiomm" onclick="fndateselect('mm');" name="joindate" value="mm" <% if joindate="mm" then %>checked<% end if %>>
			<label for="radiomm">월별</label>

			<input type="radio" id="radioww" onclick="fndateselect('ww');" name="joindate" value="ww" <% if joindate="ww" then %>checked<% end if %>>
			<label for="radioww">주차별</label>

			<input type="radio" id="radiodd" onclick="fndateselect('dd');" name="joindate" value="dd" <% if joindate="dd" then %>checked<% end if %>>
			<label for="radiodd">일별</label>
			&nbsp; &nbsp; &nbsp;
			<select name="joinyy" class="selectBox">
			<% for i = 2017 to year(now) %>
				<option value="<%=i%>" <%=chkIIF(i=year(now),"selected","")%>><%=i%></option>
			<% next %>
			</select>
			&nbsp; &nbsp; &nbsp;
			<input type="radio" id="radiosubmit" name="radiosubmit" onclick="searchFrm();">
			<label for="radiosubmit">적용</label>&nbsp; &nbsp; &nbsp;※ 실시간 데이터와 1시간정도 차이남.
			<br>
			&nbsp; &nbsp; &nbsp;
			<div id="month">
			<% for i = 1 to 12 %>
				<input type="radio" id="radio<%=i%>" name="joinmm" value="<%=i%>" <% if joinmm=i then%>checked<% end if %>>
				<label for="radio<%=i%>"><%=i%>월</label>
			<% next %>
			</div>
		</ul>

		<ul id="weeklist" style="display:none;">
			<input type="radio" id="radioweek1" name="joinww" value="1" <% if joinww = 1 then %>checked<% end if %>>
			<label for="radioweek1">1분기</label>

			<input type="radio" id="radioweek2" name="joinww" value="2" <% if joinww = 2 then %>checked<% end if %>>
			<label for="radioweek2">2분기</label>

			<input type="radio" id="radioweek3" name="joinww" value="3" <% if joinww = 3 then %>checked<% end if %>>
			<label for="radioweek3">3분기</label>

			<input type="radio" id="radioweek4" name="joinww" value="4" <% if joinww = 4 then %>checked<% end if %>>
			<label for="radioweek4">4분기</label>
		</ul>

	</div>
	</form>

	<div class="cont">
		<div class="pad5">
			<div class="tPad15">
				<table class="tbType1 listTb">
					<thead>
					<tr> 
						<th><div>구분</div></th>
						<% for i = 0 to forcnt %>
							<th><div><%= i+fordate %><%= datename %></div></th>
						<% next %>
						<th><div>합계</div></th>
					</tr>
					</thead>
					<tbody>
						<%
						dim rowcounter, colcounter, numcols, numrows, thisfield
						dim gubun
						IF isArray(alldata) THEN
							numcols=ubound(alldata,1)
							numrows=ubound(alldata,2)
							FOR rowcounter= 0 TO numrows
								response.write "<tr onMouseOver=""colorCH(this,'on')"" onMouseOut=""colorCH(this,'off')"">" & vbcrlf
								FOR colcounter=0 to numcols
									thisfield=alldata(colcounter,rowcounter)
									if isnull(thisfield) or trim(thisfield)=""then
										thisfield="0"
									end if
									response.write "<td>" 
									Select Case right(thisfield,2)
										Case "nv"
											gubun = "네이버"
										Case "ka"
											gubun = "카카오"
										Case "gl"
											gubun = "구글"
										Case "fb"
											gubun = "페이스북"
										Case "ap"
											gubun = "애플"
										case else
											gubun = ""
									End Select
									if not(rowcounter = 0 or IsNumeric(thisfield)) then
										response.write left(thisfield,3)
									else
										response.write thisfield
									end if
									if colcounter=0 then response.write gubun
									response.write "</td>" & vbcrlf
								NEXT
								response.write "</tr>" & vbcrlf
							NEXT
						else
						%>
							<tr>
								<td colspan="33">검색된 데이터가 없습니다.</td>
							</tr>
						<%
						end if
						%>
					</tbody>
				</table>
			</div>
		</div>
	</div>
</div>
</body>
<script>

<% if joindate = "dd" then %>
	$('#month').show();
	$('#weeklist').hide();
	$("#radio"+'<%=joinmm%>').attr('checked', 'checked');
<% elseif joindate = "ww" then %>
	$('#month').hide();
	$('#weeklist').show();
<% else %>
	$('#month').hide();
	$('#weeklist').hide();
<% end if %>
document.all.loadingBar.style.display='none';
</script>
</html>
<!-- 페이지 끝 -->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->
<%	response.Flush %>
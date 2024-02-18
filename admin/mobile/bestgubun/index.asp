<%@ language=vbscript %>
<% option explicit %>
<% response.Buffer=true %>
<%
'###########################################################
' Description : 모바일 BEST페이지 검색 구분, 기본값 지정
' History : 2017.11.02 유태욱
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/partner/lib/adminHead.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/partner/lib/function/incPageFunction.asp" -->


<%
dim sqlStr, bestgubun

	sqlStr = "select top 1 bestgubun " & vbcrlf
	sqlStr = sqlStr & " from db_sitemaster.dbo.tbl_mobile_best_gubun " & vbcrlf
	rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly  '' 수정.2015/08/11
	IF Not rsget.Eof Then
		bestgubun = rsget(0)
	End IF
	rsget.close

	if bestgubun="" or isNull(bestgubun) then
		bestgubun = "dt"
	end if
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
</style>

<script language="javascript">
function searchFrm(){
	alert('적용하시면 모바일 BEST페이지 구분값 기본이\n\n변경됩니다.');
	frm.submit();
}

function fndateselect(sel){
	$('#bestgubun').val(sel);
}

</script>

</head>
<body>
<div class="wrap"><br><br>
	<!-- search -->
	<form name="frm" method="get" action="bestgubun_proc.asp">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="mode" value="bestgubunupdate">
	<input type="hidden" name="bestgubun" id="bestgubun" value="<%= bestgubun %>">
	<div class="searchWrap">
		<ul>
			<input type="radio" id="radiodt" onclick="fndateselect('dt');" name="bestgubunset" value="dt" <% if bestgubun="dt" then %>checked<% end if %>>
			<label for="radiodt">기간별</label>
			&nbsp; 
			<input type="radio" id="radione" onclick="fndateselect('ne');" name="bestgubunset" value="ne" <% if bestgubun="ne" then %>checked<% end if %>>
			<label for="radione">신상품</label>
			&nbsp; 
			<input type="radio" id="radiost" onclick="fndateselect('st');" name="bestgubunset" value="st" <% if bestgubun="st" then %>checked<% end if %>>
			<label for="radiost">스테디셀러</label>
			&nbsp; 
			<input type="radio" id="radiows" onclick="fndateselect('ws');" name="bestgubunset" value="ws" <% if bestgubun="ws" then %>checked<% end if %>>
			<label for="radiows">위시베스트</label>
			&nbsp; 
			<input type="radio" id="radioag" onclick="fndateselect('ag');" name="bestgubunset" value="ag" <% if bestgubun="ag" then %>checked<% end if %>>
			<label for="radioag">연령별</label>
			&nbsp; 
			<input type="radio" id="radiolt" onclick="fndateselect('lt');" name="bestgubunset" value="lt" <% if bestgubun="lt" then %>checked<% end if %>>
			<label for="radiolt">후기</label>
			&nbsp; 
			<input type="radio" id="radiolv" onclick="fndateselect('lv');" name="bestgubunset" value="lv" <% if bestgubun="lv" then %>checked<% end if %>>
			<label for="radiolv">등급별</label>
			&nbsp; 
			<input type="radio" id="radiobr" onclick="fndateselect('br');" name="bestgubunset" value="br" <% if bestgubun="br" then %>checked<% end if %>>
			<label for="radiobr">브랜드</label>
			&nbsp; 
			<input type="radio" id="radioms" onclick="fndateselect('ms');" name="bestgubunset" value="ms" <% if bestgubun="ms" then %>checked<% end if %>>
			<label for="radioms">맨즈</label>
			&nbsp; 
			<input type="radio" id="radiofr" onclick="fndateselect('fr');" name="bestgubunset" value="fr" <% if bestgubun="fr" then %>checked<% end if %>>
			<label for="radiofr">첫구매</label>
			&nbsp; &nbsp; &nbsp;&nbsp; &nbsp; &nbsp;

			<input type="radio" id="radiosubmit" name="radiosubmit" onclick="searchFrm();">
			<label for="radiosubmit">적용</label>
			<br>
		</ul>
	</div>
	</form>
</div>
</body>
</html>
<!-- #include virtual="/lib/db/dbclose.asp" -->

<%@ language=vbscript %>
<% option explicit %>
<%
'###############################################
' PageName : play_insert.asp
' Discription : 모바일 play
' History : 2015-05-13 이종화
'###############################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/classes/mobile/main_play.asp" -->
<%
Dim idx , subImage1 , isusing , mode
Dim srcSDT , srcEDT 
Dim mainStartDate, mainEndDate
Dim sDt, sTm, eDt, eTm
Dim sortnum  , prevDate , ordertext
Dim title , gubun , subtitle , url_mo , url_app , appdiv , appcate
Dim itemid , itemName , smallImage
	idx = requestCheckvar(request("idx"),16)
	srcSDT = request("sDt")
	srcEDT = request("eDt")
	prevDate = request("prevDate")

If idx = "" Then 
	mode = "add" 
Else 
	mode = "modify" 
End If 

If idx <> "" then
	dim playOne
	set playOne = new CMainbanner
	playOne.FRectIdx = idx
	playOne.GetOneContents()

	gubun				=	playOne.FOneItem.Fgubun
	title				=	playOne.FOneItem.Ftitle
	subtitle			=	playOne.FOneItem.Fsubtitle
	url_mo				=	playOne.FOneItem.Furl_mo
	url_app				=	playOne.FOneItem.Furl_app
	appdiv				=	playOne.FOneItem.Fappdiv
	appcate				=	playOne.FOneItem.fappcate
	sortnum				=	playOne.FOneItem.Fsortnum
	mainStartDate		=	playOne.FOneItem.Fstartdate
	mainEndDate			=	playOne.FOneItem.Fenddate 
	isusing				=	playOne.FOneItem.Fisusing
	subImage1			=	playOne.FOneItem.Fpimg
	ordertext			=	playOne.FOneItem.Fordertext

	set playOne = Nothing
End If 

if Not(mainStartDate="" or isNull(mainStartDate)) then
	sDt = left(mainStartDate,10)
	sTm = Num2Str(hour(mainStartDate),2,"0","R") &":"& Num2Str(minute(mainStartDate),2,"0","R") &":"& Num2Str(second(mainStartDate),2,"0","R")
else
	if srcSDT<>"" then
		sDt = left(srcSDT,10)
	else
		sDt = date
	end if
	sTm = "00:00:00"
end if

if Not(mainEndDate="" or isNull(mainEndDate)) then
	eDt = left(mainEndDate,10)
	eTm = Num2Str(hour(mainEndDate),2,"0","R") &":"& Num2Str(minute(mainEndDate),2,"0","R") &":"& Num2Str(second(mainEndDate),2,"0","R")
else
	if srcEDT<>"" then
		eDt = left(srcEDT,10)
	else
		eDt = date
	end if
	eTm = "23:59:59"
end if
%>
<link href="/js/jqueryui/css/jquery-ui.css" rel="stylesheet">
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/jqueryui/jquery-ui-1.10.2.custom.min.js"></script>
<script type='text/javascript'>
	function jsSubmit(){
		var frm = document.frm;

		if (frm.gubun.value == 0 )
		{
			alert('구분을 선택 하세요');
			frm.gubun.focus();
			return;
		}

//		if (!frm.title.value)
//		{
//			alert('제목을 입력해주세요.');
//			frm.title.focus();
//			return;
//		}
//
//		if (!frm.subtitle.value)
//		{
//			alert('내용을 입력해주세요.');
//			frm.subtitle.focus();
//			return;
//		}

		if (!frm.url_mo.value)
		{
			alert('모바일 URL을 입력해주세요.');
			frm.url_mo.focus();
			return;
		}

		if (!frm.url_app.value)
		{
			alert('앱 URL을 입력해주세요.');
			frm.url_app.focus();
			return;
		}
	
		if (confirm('저장 하시겠습니까?')){
			//frm.target = "blank";
			frm.submit();
		}
	}
	function jsgolist(){
		self.location.href="/admin/mobile/play/";
	}
	$(function(){
	//달력대화창 설정
	var arrDayMin = ["일","월","화","수","목","금","토"];
	var arrMonth = ["1월","2월","3월","4월","5월","6월","7월","8월","9월","10월","11월","12월"];
    $("#sDt").datepicker({
		dateFormat: "yy-mm-dd",
		prevText: '이전달', nextText: '다음달', yearSuffix: '년',
		dayNamesMin: arrDayMin,
		monthNames: arrMonth,
		showMonthAfterYear: true,
    	numberOfMonths: 2,
    	showCurrentAtPos: 1,
      	showOn: "button",
      	<% if Idx<>"" then %>maxDate: "<%=eDt%>",<% end if %>
    	onClose: function( selectedDate ) {
    		$( "#eDt" ).datepicker( "option", "minDate", selectedDate );
    	}
    });
    $("#eDt").datepicker({
		dateFormat: "yy-mm-dd",
		prevText: '이전달', nextText: '다음달', yearSuffix: '년',
		dayNamesMin: arrDayMin,
		monthNames: arrMonth,
		showMonthAfterYear: true,
    	numberOfMonths: 2,
      	showOn: "button",
      	<% if Idx<>"" then %>minDate: "<%=sDt%>",<% end if %>
    	onClose: function( selectedDate ) {
    		$( "#sDt" ).datepicker( "option", "maxDate", selectedDate );
    	}
    });
});

function putLinkText(key,gubun) {
	var frm = document.frm;
	var urllink
	if (gubun == "1" )
	{
		urllink = frm.url_mo;
	}

	switch(key) {
		case 'search':
			urllink.value='/search/search_result.asp?rect=검색어';
			break;
		case 'event':
			urllink.value='/event/eventmain.asp?eventid=이벤트번호';
			break;
		case 'itemid':
			urllink.value='/category/category_itemprd.asp?itemid=상품코드';
			break;
		case 'category':
			urllink.value='/category/category_list.asp?disp=카테고리';
			break;
		case 'brand':
			urllink.value='/street/street_brand.asp?makerid=브랜드아이디';
			break;
		case 'ground':
			urllink.value='/play/playGround.asp?idx=그라운드번호&contentsidx=컨텐츠번호';
			break;
		case 'style+':
			urllink.value='/play/playStylePlus.asp?idx=스타일플러스번호&contentsidx=컨텐츠번호';
			break;
		case 'fingers':
			urllink.value='/play/playDesignFingers.asp?idx=핑거스번호&contentsidx=컨텐츠번호';
			break;
		case 't-episode':
			urllink.value='/play/playTEpisode.asp?idx=티에피소드번호&contentsidx=컨텐츠번호';
			break;
		case 'gift':
			urllink.value='/gift/gifttalk/';
			break;
	}
}

//url 자동 생성
function chklink(v){
	if (v == "1"){
		document.frm.url_app.value = "/apps/appCom/wish/web2014/category/category_itemprd.asp?itemid=상품코드";
		$("#catesel").css("display","none");
		$("#url_app").prop('disabled',false);
	}else if (v == "2"){
		document.frm.url_app.value = "/apps/appcom/wish/web2014/event/eventmain.asp?eventid=이벤트코드&rdsite=rdsite명(필수아님)";
		$("#catesel").css("display","none");
		$("#url_app").prop('disabled',false);
	}else if (v == "3"){
		document.frm.url_app.value = "makerid=브랜드명";
		$("#catesel").css("display","none");
		$("#url_app").prop('disabled',false);
	}else if (v == "4"){
		chgDispCate2('');
		document.frm.url_app.value = "cd1=&nm1=";
		$("#catesel").css("display","block");
		$("#url_app").attr('readonly','readonly');
	}else if (v == "5"){//'ground
		document.frm.url_app.value = "/apps/appcom/wish/web2014/play/playGround.asp?idx=그라운드번호&contentsidx=컨텐츠번호";
		$("#catesel").css("display","none");
		$("#url_app").prop('disabled',false);
	}else if (v == "6"){//'style+
		document.frm.url_app.value = "/apps/appcom/wish/web2014/play/playStylePlus.asp?idx=스타일플러스번호&contentsidx=컨텐츠번호";
		$("#catesel").css("display","none");
		$("#url_app").prop('disabled',false);
	}else if (v == "7"){//'fingers
		document.frm.url_app.value = "/apps/appcom/wish/web2014/play/playDesignFingers.asp?idx=핑거스번호&contentsidx=컨텐츠번호";
		$("#catesel").css("display","none");
		$("#url_app").prop('disabled',false);
	}else if (v == "8"){//'t-episode
		document.frm.url_app.value = "/apps/appcom/wish/web2014/play/playTEpisode.asp?idx=티에피소드번호&contentsidx=컨텐츠번호";
		$("#catesel").css("display","none");
		$("#url_app").prop('disabled',false);
	}else if (v == "9"){//'gift
		document.frm.url_app.value = "/apps/appcom/wish/web2014/gift/gifttalk/";
		$("#catesel").css("display","none");
		$("#url_app").prop('disabled',false);
	}else{
		document.frm.url_app.value = "APP URL 구분을 선택 해주세요.";
		$("#catesel").css("display","none");
		$("#url_app").prop('disabled',false);
	}
}

function chgDispCate2(dc) {
	$.ajax({
		url: "/admin/mobile/catetag/dispCateSelectBox_response.asp?disp="+dc,
		cache: false,
		async: false,
		success: function(message) {
			// 내용 넣기
			$("#lyrDispCtBox2").empty().html(message);
			if (dc.length == 3){
				document.frm.url_app.value = $("#dispcateval1 option:selected").val()+"||"+$("#dispcateval1 option:selected").text();
				$("#appcate").val(dc);
			}else if (dc.length == 6){
				document.frm.url_app.value = $("#dispcateval1 option:selected").val()+"||"+$("#dispcateval2 option:selected").val()+"||"+$("#dispcateval1 option:selected").text()+"||"+$("#dispcateval2 option:selected").text();
				$("#appcate").val(dc);
			}else if (dc.length == 9){
				document.frm.url_app.value = $("#dispcateval1 option:selected").val()+"||"+$("#dispcateval2 option:selected").val()+"||"+$("#dispcateval3 option:selected").val()+"||"+$("#dispcateval1 option:selected").text()+"||"+$("#dispcateval2 option:selected").text()+"||"+$("#dispcateval3 option:selected").text();
				$("#appcate").val(dc);
			}else{
				
			}

		}
	});
}
$(function(){
	chgDispCate2('<%=appcate%>');
});
</script>
<table width="900" cellpadding="2" cellspacing="1" class="a" bgcolor="#3d3d3d">
<form name="frm" method="post" action="<%=uploadUrl%>/linkweb/mobile/playbanner_proc.asp" enctype="multipart/form-data" style="margin:0px;">
<input type="hidden" name="mode" value="<%=mode%>">
<input type="hidden" name="adminid" value="<%=session("ssBctId")%>">
<input type="hidden" name="idx" value="<%=idx%>">
<input type="hidden" name="prevDate" value="<%=prevDate%>">
<input type="hidden" name="appcate" id="appcate"/>
<tr bgcolor="#FFFFFF">
    <td bgcolor="#FFF999" align="center" width="20">노출기간</td>
    <td >
		<input type="text" id="sDt" name="StartDate" size="10" value="<%=sDt%>" />
		<input type="text" name="sTm" size="8" value="<%=sTm%>" /> ~
		<input type="text" id="eDt" name="EndDate" size="10" value="<%=eDt%>" />
		<input type="text" name="eTm" size="8" value="<%=eTm%>" />
    </td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center">구분-text입력<br/>ex)GROUND,Fingers,Gift etc...</td>
	<td >
		<select name="gubun">
			<option value="0">선택하세요</option>
			<option value="1" <% if gubun = "1" then response.write " selected" %>>GROUND</option>
			<option value="2" <% if gubun = "2" then response.write " selected" %>>STYLE+</option>
			<option value="3" <% if gubun = "3" then response.write " selected" %>>DESIGN FINGERS</option>
			<option value="4" <% if gubun = "4" then response.write " selected" %>>T-EPISODE</option>
			<option value="5" <% if gubun = "5" then response.write " selected" %>>GIFT</option>
		</select>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center">타이틀</td>
	<td ><input type="text" name="title" size="50" value="<%=title%>"/></td>
</tr>
<!-- <tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center">서브타이틀</td>
	<td ><input type="text" name="subtitle" size="50" value="<%=subtitle%>"/></td>
</tr> -->
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center" >이미지</td>
	<td>
		<input type="file" name="subImage1" class="file" title="이미지 #1" require="N" style="width:80%;" />
		<% if subImage1<>"" then %>
		<br>
		<img src="<%= subImage1 %>" width="100" /><br><%= subImage1 %>
		[<span style="color:red">이미지삭제</span>] --&gt; <input type="checkbox" name="delimg" value="1"/>
		<% end if %>		
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center">모바일 URL</td>
	<td ><input type="text" name="url_mo" size="80" value="<%=url_mo%>"/>
	<br/><br/>ex)
		<font color="#707070">
		- <span style="cursor:pointer" onClick="putLinkText('event','1')">이벤트 링크 : /event/eventmain.asp?eventid=<font color="darkred">이벤트코드</font></span><br>
		- <span style="cursor:pointer" onClick="putLinkText('itemid','1')">상품코드 링크 : /category/category_itemprd.asp?itemid=<font color="darkred">상품코드 (O)</font></span><br>
		- <span style="cursor:pointer" onClick="putLinkText('category','1')">카테고리 링크 : /category/category_list.asp?disp=<font color="darkred">카테고리</font></span><br>
		- <span style="cursor:pointer" onClick="putLinkText('brand','1')">브랜드아이디 링크 : /street/street_brand.asp?makerid=<font color="darkred">브랜드아이디</font></span><br>
		- <span style="cursor:pointer" onClick="putLinkText('ground','1')">그라운드 링크 : /play/playGround.asp?idx=<font color="darkred">그라운드번호</font>&contentsidx=<font color="darkred">컨텐츠번호</font></span><br>
		- <span style="cursor:pointer" onClick="putLinkText('style+','1')">스타일플러스 링크 : /play/playStylePlus.asp?idx=<font color="darkred">스타일플러스번호</font>&contentsidx=<font color="darkred">컨텐츠번호</font></span><br>
		- <span style="cursor:pointer" onClick="putLinkText('fingers','1')">디자인핑거스 링크 : /play/playDesignFingers.asp?idx=<font color="darkred">핑거스번호</font>&contentsidx=<font color="darkred">컨텐츠번호</font></span><br>
		- <span style="cursor:pointer" onClick="putLinkText('t-episode','1')">티에피소드 링크 : /play/playTEpisode.asp?idx=<font color="darkred">티에피소드번호</font>&contentsidx=<font color="darkred">컨텐츠번호</font></span><br>
		- <span style="cursor:pointer" onClick="putLinkText('gift','1')">기프트 링크 : <font color="darkred">/gift/gifttalk/</font></span>
		</font>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center">앱 URL</td>
	<td >
		<table width="100%" cellpadding="3" cellspacing="1" class="a" bgcolor="#3d3d3d">
			<tr>
				<td bgcolor="#FFF999" width="100" align="center">APP URL 구분</td>
				<td bgcolor="#FFFFFF">
					<select name='appdiv' class='select' onchange="chklink(this.value);">
						<option value="0">선택하세요</option>
						<option value="1" <% if appdiv = "1" then response.write " selected" %>>상품상세</option>
						<option value="2" <% if appdiv = "2" then response.write " selected" %>>이벤트</option>
						<option value="3" <% if appdiv = "3" then response.write " selected" %>>브랜드</option>
						<option value="4" <% if appdiv = "4" then response.write " selected" %>>카테고리</option>
						<option value="5" <% if appdiv = "5" then response.write " selected" %>>Ground</option>
						<option value="6" <% if appdiv = "6" then response.write " selected" %>>Style+</option>
						<option value="7" <% if appdiv = "7" then response.write " selected" %>>Fingers</option>
						<option value="8" <% if appdiv = "8" then response.write " selected" %>>T-episode</option>
						<option value="9" <% if appdiv = "9" then response.write " selected" %>>기프트</option>
					</select>
				</td>
			</tr>
			<tr id="catesel" style="display:<%=chkiif(idx<>"" And appdiv = "4","block","none")%>">
				<td bgcolor="#FFF999" width="100" align="center">전시카테고리 선택</td>
				<td bgcolor="#FFFFFF">
					<span id="lyrDispCtBox2"></span>
				</td>
			</tr>
			<tr>
				<td bgcolor="#FFF999" width="100" align="center">코드내용</td>
				<td bgcolor="#FFFFFF"><textarea name="url_app" class="textarea" id="url_app" style="width:100%; height:40px;"><%=url_app%></textarea></td>
			</tr>
		</table>
	</td>
</tr>

<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center">사용여부</td>
	<td ><div style="float:left;"><input type="radio" name="isusing" value="Y" <%=chkiif(isusing = "Y","checked","")%> checked />사용함 &nbsp;&nbsp;&nbsp; <input type="radio" name="isusing" value="N"  <%=chkiif(isusing = "N","checked","")%>/>사용안함</div> <div style="float:right;margin-top:5px;margin-right:10px;"></div></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center">정렬번호</td>
	<td ><input type="text" name="sortnum" value="<%=chkiif(sortnum="","0",sortnum)%>" size="2"/></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center">작업자 지시사항</td>
	<td ><textarea name="ordertext" cols="80" rows="8"/><%=ordertext%></textarea></td>
</tr>
<tr bgcolor="#FFFFFF" align="center">
    <td colspan="2"><input type="button" value=" 취 소 " onClick="jsgolist();"/><input type="button" value=" 저 장 " onClick="jsSubmit();"/></td>
</tr>
</form>
</table>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
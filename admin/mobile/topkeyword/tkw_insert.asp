<%@ language=vbscript %>
<% option explicit %>
<%
'###############################################
' PageName : tkw_insert.asp
' Discription : 모바일 GNB top keyword
' History : 2015-09-16 이종화
'###############################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/classes/mobile/topkeyword.asp" -->
<!-- #include virtual="/lib/classes/mobile/topcatecodeCls.asp" -->
<%
Dim idx , subImage1 , isusing , mode , gcode
Dim srcSDT , srcEDT 
Dim mainStartDate, mainEndDate
Dim sDt, sTm, eDt, eTm
Dim sortnum  , prevDate , ordertext
Dim ktitle , kword , kcontents , kurl_mo , kurl_app , appdiv , appcate
Dim itemid , itemName , smallImage
	idx = requestCheckvar(request("idx"),16)
	gcode = requestCheckvar(request("gcode"),3)
	srcSDT = request("sDt")
	srcEDT = request("eDt")
	prevDate = request("prevDate")

If idx = "" Then 
	mode = "add" 
Else 
	mode = "modify" 
End If 

If idx <> "" then
	dim tkwobannerOne
	set tkwobannerOne = new CMainbanner
	tkwobannerOne.FRectIdx = idx
	tkwobannerOne.GetOneContents()

	kword				=	tkwobannerOne.FOneItem.Fkword
	ktitle				=	tkwobannerOne.FOneItem.Fktitle
	kcontents			=	tkwobannerOne.FOneItem.Fkcontents
	kurl_mo				=	tkwobannerOne.FOneItem.Fkurl_mo
	kurl_app			=	tkwobannerOne.FOneItem.Fkurl_app
	appdiv				=	tkwobannerOne.FOneItem.Fappdiv
	appcate				=	tkwobannerOne.FOneItem.fappcate
	sortnum				=	tkwobannerOne.FOneItem.Fsortnum
	mainStartDate		=	tkwobannerOne.FOneItem.Fstartdate
	mainEndDate			=	tkwobannerOne.FOneItem.Fenddate 
	isusing				=	tkwobannerOne.FOneItem.Fisusing
	subImage1			=	tkwobannerOne.FOneItem.Fkwimg
	ordertext			=	tkwobannerOne.FOneItem.Fordertext
	itemid				=	tkwobannerOne.FOneItem.Fitemid
	itemname			=	tkwobannerOne.FOneItem.Fitemname
	smallimage			=	tkwobannerOne.FOneItem.Fsmallimage
	gcode				=	tkwobannerOne.FOneItem.Fgnbcode

	set tkwobannerOne = Nothing
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

		if (!frm.gcode.value)
		{
			alert('노출 GNB 영역을 선택 해주세요.');
			frm.gcode.focus();
			return;
		}

//		if (!frm.kword.value)
//		{
//			alert('키워드를 입력해주세요.');
//			frm.kword.focus();
//			return;
//		}

		if (!frm.ktitle.value)
		{
			alert('제목을 입력해주세요.');
			frm.ktitle.focus();
			return;
		}
//
//		if (!frm.kcontents.value)
//		{
//			alert('내용을 입력해주세요.');
//			frm.kcontents.focus();
//			return;
//		}

		if (!frm.kurl_mo.value)
		{
			alert('모바일 URL을 입력해주세요.');
			frm.kurl_mo.focus();
			return;
		}

		if (!frm.kurl_app.value)
		{
			alert('앱 URL을 입력해주세요.');
			frm.kurl_app.focus();
			return;
		}
	
		if (confirm('저장 하시겠습니까?')){
			//frm.target = "blank";
			frm.submit();
		}
	}
	function jsgolist(){
		self.location.href="/admin/mobile/topkeyword/";
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
    	numberOfMonths: 1,
    	showCurrentAtPos: 0,
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
		urllink = frm.kurl_mo;
	}

	switch(key) {
//		case 'search':
//			urllink.value='/search/search_result.asp?rect=검색어';
//			break;
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
	}
}

//url 자동 생성
function chklink(v){
	if (v == "1"){
		document.frm.kurl_app.value = "/apps/appCom/wish/web2014/category/category_itemprd.asp?itemid=상품코드";
		alert('Mobile URL 복사 금지!');
		$("#catesel").css("display","none");
		$("#kurl_app").prop('disabled',false);
	}else if (v == "2"){
		document.frm.kurl_app.value = "/apps/appcom/wish/web2014/event/eventmain.asp?eventid=이벤트코드&rdsite=rdsite명(필수아님)";
		alert('Mobile URL 복사 금지!');
		$("#catesel").css("display","none");
		$("#kurl_app").prop('disabled',false);
	}else if (v == "3"){
		document.frm.kurl_app.value = "makerid=브랜드명";
		alert('Mobile URL 복사 금지!');
		$("#catesel").css("display","none");
		$("#kurl_app").prop('disabled',false);
	}else if (v == "4"){
		chgDispCate2('');
		document.frm.kurl_app.value = "cd1=&nm1=";
		$("#catesel").css("display","block");
		$("#kurl_app").attr('readonly','readonly');
	}else{
		document.frm.kurl_app.value = "APP URL 구분을 선택 해주세요.";
		$("#catesel").css("display","none");
		$("#kurl_app").prop('disabled',false);
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
				document.frm.kurl_app.value = $("#dispcateval1 option:selected").val()+"||"+$("#dispcateval1 option:selected").text();
				$("#appcate").val(dc);
			}else if (dc.length == 6){
				document.frm.kurl_app.value = $("#dispcateval1 option:selected").val()+"||"+$("#dispcateval2 option:selected").val()+"||"+$("#dispcateval1 option:selected").text()+"||"+$("#dispcateval2 option:selected").text();
				$("#appcate").val(dc);
			}else if (dc.length == 9){
				document.frm.kurl_app.value = $("#dispcateval1 option:selected").val()+"||"+$("#dispcateval2 option:selected").val()+"||"+$("#dispcateval3 option:selected").val()+"||"+$("#dispcateval1 option:selected").text()+"||"+$("#dispcateval2 option:selected").text()+"||"+$("#dispcateval3 option:selected").text();
				$("#appcate").val(dc);
			}else{
				
			}

		}
	});
}
$(function(){
	<% if appdiv ="4" then %>
	chgDispCate2('<%=appcate%>');
	<% end if %>
});

// 상품정보 접수
function fnGetItemInfo(iid) {
	$.ajax({
		type: "GET",
		url: "/admin/sitemaster/wcms/act_iteminfo.asp?itemid="+iid,
		dataType: "xml",
		cache: false,
		async: false,
		timeout: 5000,
		beforeSend: function(x) {
			if(x && x.overrideMimeType) {
				x.overrideMimeType("text/xml;charset=euc-kr");
			}
		},
		success: function(xml) {
			if($(xml).find("itemInfo").find("item").length>0) {
				var rst = "<img src='" + $(xml).find("itemInfo").find("item").find("smallImage").text() + "' height='50' /><br/>"
					rst += $(xml).find("itemInfo").find("item").find("itemname").text();
				$("#lyItemInfo").fadeIn();
				$("#lyItemInfo").html(rst);
			} else {
				$("#lyItemInfo").fadeOut();
			}
		},
		error: function(xhr, status, error) {
			$("#lyItemInfo").fadeOut();
			/*alert(xhr + '\n' + status + '\n' + error);*/
		}
	});
}

</script>
<table width="900" cellpadding="2" cellspacing="1" class="a" bgcolor="#3d3d3d">
<form name="frm" method="post" action="<%=uploadUrl%>/linkweb/mobile/tkwbanner_proc.asp" enctype="multipart/form-data" style="margin:0px;">
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
	<td bgcolor="#FFF999" align="center">노출 GNB영역</td>
	<td ><% Call drawSelectBoxGNB("gcode" , gcode) %></td>
</tr>
<!-- <tr bgcolor="#FFFFFF"> -->
<!-- 	<td bgcolor="#FFF999" align="center">키워드</td> -->
<!-- 	<td ><input type="text" name="kword" size="50" value="<%=kword%>"/></td> -->
<!-- </tr> -->
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center">제목</td>
	<td ><input type="text" name="ktitle" size="50" value="<%=ktitle%>"/></td>
</tr>
<!-- <tr bgcolor="#FFFFFF"> -->
<!-- 	<td bgcolor="#FFF999" align="center">내용</td> -->
<!-- 	<td ><textarea name="kcontents" cols="50" rows="4"/><%=kcontents%></textarea></td> -->
<!-- </tr> -->
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center" colspan="2">상품코드와 이미지중 <span style="color:red">한가지 이상</span> 입력 - 이미지가 우선으로 뿌려집니다.</td>
</tr>
<tr bgcolor="#FFFFFF">
    <td bgcolor="#FFF999" align="center">상품코드</td>
    <td colspan="3">
        <input type="text" name="itemid" value="<%= itemid %>" size="8" maxlength="8" class="text" require="N" onblur="fnGetItemInfo(this.value)" title="상품코드" />
        <div id="lyItemInfo" style="display:<%=chkIIF(itemid="","none","")%>;">
        <%
        	if Not(itemName="" or isNull(itemName)) then
        		Response.Write "<img src='" & smallImage & "' height='50' /><br/>"
        		Response.Write itemName
        	end if
        %>
        </div>
    </td>
</tr>
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
	<td ><input type="text" name="kurl_mo" size="80" value="<%=kurl_mo%>"/>
	<br/><br/>ex)
		<font color="#707070">
		<!-- - <span style="cursor:pointer" onClick="putLinkText('search','1')">검색결과 링크 : /search/search_result.asp?rect=<font color="darkred">검색어</font></span><br> -->
		- <span style="cursor:pointer" onClick="putLinkText('event','1')">이벤트 링크 : /event/eventmain.asp?eventid=<font color="darkred">이벤트코드</font></span><br>
		- <span style="cursor:pointer" onClick="putLinkText('itemid','1')">상품코드 링크 : /category/category_itemprd.asp?itemid=<font color="darkred">상품코드 (O)</font></span><br>
		- <span style="cursor:pointer" onClick="putLinkText('category','1')">카테고리 링크 : /category/category_list.asp?disp=<font color="darkred">카테고리</font></span><br>
		- <span style="cursor:pointer" onClick="putLinkText('brand','1')">브랜드아이디 링크 : /street/street_brand.asp?makerid=<font color="darkred">브랜드아이디</font></span>
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
				<td bgcolor="#FFFFFF"><textarea name="kurl_app" class="textarea" id="kurl_app" style="width:100%; height:40px;"><%=kurl_app%></textarea></td>
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
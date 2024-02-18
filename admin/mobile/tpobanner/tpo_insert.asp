<%@ language=vbscript %>
<% option explicit %>
<%
'###############################################
' PageName : tpo_insert.asp
' Discription : 모바일 tpobanner
' History : 2013.12.14 이종화
'###############################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/classes/mobile/tpobanner.asp" -->
<%
Dim idx , subImage1 , subImage2 , subImage3 , isusing , mode
Dim srcSDT , srcEDT 
Dim mainStartDate, mainEndDate
Dim sDt, sTm, eDt, eTm
Dim lalt , ralt , lurl , rurl , sortnum  , prevDate , ordertext
Dim bgalt , bgurl
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
	dim otpobannerOne
	set otpobannerOne = new CMainbanner
	otpobannerOne.FRectIdx = idx
	otpobannerOne.GetOneContents()

	lalt				=	otpobannerOne.FOneItem.Flalt
	ralt				=	otpobannerOne.FOneItem.Fralt
	bgalt				=	otpobannerOne.FOneItem.Fbgalt
	lurl				=	otpobannerOne.FOneItem.Flurl
	rurl				=	otpobannerOne.FOneItem.Frurl
	bgurl				=	otpobannerOne.FOneItem.Fbgurl
	sortnum		=	otpobannerOne.FOneItem.Fsortnum
	mainStartDate	=	otpobannerOne.FOneItem.Fstartdate
	mainEndDate	=	otpobannerOne.FOneItem.Fenddate 
	isusing			=	otpobannerOne.FOneItem.Fisusing
	subImage1	=	otpobannerOne.FOneItem.Fbgimg
	subImage2	=	otpobannerOne.FOneItem.Flimg
	subImage3	=	otpobannerOne.FOneItem.Frimg
	ordertext		=	otpobannerOne.FOneItem.Fordertext

	set otpobannerOne = Nothing
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
	
		if (!frm.ltalt.value)
		{
			alert('좌측배너 alt를 선택해주세요');
			frm.ltalt.focus();
			return;
		}

		if (!frm.rtalt.value)
		{
			alert('우측배너 alt를 선택해주세요');
			frm.rtalt.focus();
			return;
		}

		if (!frm.lturl.value)
		{
			alert('좌측배너 url를 선택해주세요');
			frm.lturl.focus();
			return;
		}

		if (!frm.rturl.value)
		{
			alert('우측배너 url를 선택해주세요');
			frm.rturl.focus();
			return;
		}

		if (confirm('저장 하시겠습니까?')){
			//frm.target = "blank";
			frm.submit();
		}
	}
	function jsgolist(){
		self.location.href="/admin/mobile/tpobanner/";
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
		urllink = frm.lturl;
	}else if (gubun == "2" ){
		urllink = frm.rturl;
	}else{
		urllink = frm.bgurl;
	}
	switch(key) {
		case 'search':
			urllink.value='/search/search_item.asp?rect=검색어';
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
	}
}
</script>
<table width="900" cellpadding="2" cellspacing="1" class="a" bgcolor="#3d3d3d">
<form name="frm" method="post" action="<%=uploadUrl%>/linkweb/mobile/tpobanner_proc.asp" enctype="multipart/form-data" style="margin:0px;">
<input type="hidden" name="mode" value="<%=mode%>">
<input type="hidden" name="adminid" value="<%=session("ssBctId")%>">
<input type="hidden" name="idx" value="<%=idx%>">
<input type="hidden" name="prevDate" value="<%=prevDate%>">
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center">작업자 지시사항</td>
	<td colspan="3"><textarea name="ordertext" cols="80" rows="8"/><%=ordertext%></textarea></td>
</tr>
<tr bgcolor="#FFFFFF">
    <td bgcolor="#FFF999" align="center" width="15%">노출기간</td>
    <td colspan="3">
		<input type="text" id="sDt" name="StartDate" size="10" value="<%=sDt%>" />
		<input type="text" name="sTm" size="8" value="<%=sTm%>" /> ~
		<input type="text" id="eDt" name="EndDate" size="10" value="<%=eDt%>" />
		<input type="text" name="eTm" size="8" value="<%=eTm%>" />
    </td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center" width="15%">BG이미지</td>
	<td width="45%">
		<input type="file" name="subImage1" class="file" title="이미지 #1" require="N" style="width:80%;" />
		<% if subImage1<>"" then %>
		<br>
		<img src="<%= subImage1 %>" width="100" /><br><%= subImage1 %>
		<% end if %>		
	</td>
	<td bgcolor="#FFF999" align="center" width="10%">BG이미지 alt</td>
	<td width="20%"><input type="text" name="bgalt" value="<%=bgalt%>" size="20" maxlength="20"/></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center">BG이미지 URL</td>
	<td colspan="3"><input type="text" name="bgurl" size="80" value="<%=bgurl%>"/>
	<br/><br/>ex)
		<font color="#707070">
		- <span style="cursor:pointer" onClick="putLinkText('search','3')">검색결과 링크 : /search/search_item.asp?rect=<font color="darkred">검색어</font></span><br>
		- <span style="cursor:pointer" onClick="putLinkText('event','3')">이벤트 링크 : /event/eventmain.asp?eventid=<font color="darkred">이벤트코드</font></span><br>
		- <span style="cursor:pointer" onClick="putLinkText('itemid','3')">상품코드 링크 : /category/category_itemprd.asp?itemid=<font color="darkred">상품코드 (O)</font></span><br>
		- <span style="cursor:pointer" onClick="putLinkText('category','3')">카테고리 링크 : /category/category_list.asp?cdl=<font color="darkred">카테고리</font></span><br>
		- <span style="cursor:pointer" onClick="putLinkText('brand','3')">브랜드아이디 링크 : /street/street_brand.asp?makerid=<font color="darkred">브랜드아이디</font></span>
		</font>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center" width="15%">좌측배너</td>
	<td width="45%">
		<input type="file" name="subImage2" class="file" title="이미지 #1" require="N" style="width:80%;" />
		<% if subImage2<>"" then %>
		<br>
		<img src="<%= subImage2 %>" width="100" /><br><%= subImage2 %>
		<% end if %>		
	</td>
	<td bgcolor="#FFF999" align="center" width="10%">좌측배너 alt</td>
	<td width="20%"><input type="text" name="ltalt" value="<%=lalt%>" size="40" maxlength="40"/></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center">좌측배너 URL</td>
	<td colspan="3"><input type="text" name="lturl" size="80" value="<%=lurl%>"/>
	<br/><br/>ex)
		<font color="#707070">
		- <span style="cursor:pointer" onClick="putLinkText('search','1')">검색결과 링크 : /search/search_item.asp?rect=<font color="darkred">검색어</font></span><br>
		- <span style="cursor:pointer" onClick="putLinkText('event','1')">이벤트 링크 : /event/eventmain.asp?eventid=<font color="darkred">이벤트코드</font></span><br>
		- <span style="cursor:pointer" onClick="putLinkText('itemid','1')">상품코드 링크 : /category/category_itemprd.asp?itemid=<font color="darkred">상품코드 (O)</font></span><br>
		- <span style="cursor:pointer" onClick="putLinkText('category','1')">카테고리 링크 : /category/category_list.asp?disp=<font color="darkred">카테고리</font></span><br>
		- <span style="cursor:pointer" onClick="putLinkText('brand','1')">브랜드아이디 링크 : /street/street_brand.asp?makerid=<font color="darkred">브랜드아이디</font></span>
		</font>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center">우측배너</td>
	<td>
		<input type="file" name="subImage3" class="file" title="이미지 #1" require="N" style="width:80%;" />
		<% if subImage3<>"" then %>
		<br>
		<img src="<%= subImage3 %>" width="100" /><br><%= subImage3 %>
		<% end if %>		
	</td>
	<td bgcolor="#FFF999" align="center">우측배너 alt</td>
	<td><input type="text" name="rtalt" value="<%=ralt%>" size="40" maxlength="40"/></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center">우측배너 URL</td>
	<td colspan="3"><input type="text" name="rturl" size="80" value="<%=rurl%>"/>
	<br/><br/>ex)
		<font color="#707070">
		- <span style="cursor:pointer" onClick="putLinkText('search','2')">검색결과 링크 : /search/search_item.asp?rect=<font color="darkred">검색어</font></span><br>
		- <span style="cursor:pointer" onClick="putLinkText('event','2')">이벤트 링크 : /event/eventmain.asp?eventid=<font color="darkred">이벤트코드</font></span><br>
		- <span style="cursor:pointer" onClick="putLinkText('itemid','2')">상품코드 링크 : /category/category_itemprd.asp?itemid=<font color="darkred">상품코드 (O)</font></span><br>
		- <span style="cursor:pointer" onClick="putLinkText('category','2')">카테고리 링크 : /category/category_list.asp?disp=<font color="darkred">카테고리</font></span><br>
		- <span style="cursor:pointer" onClick="putLinkText('brand','2')">브랜드아이디 링크 : /street/street_brand.asp?makerid=<font color="darkred">브랜드아이디</font></span>
		</font>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center">사용여부</td>
	<td colspan="3"><div style="float:left;"><input type="radio" name="isusing" value="Y" <%=chkiif(isusing = "Y","checked","")%> checked />사용함 &nbsp;&nbsp;&nbsp; <input type="radio" name="isusing" value="N"  <%=chkiif(isusing = "N","checked","")%>/>사용안함</div> <div style="float:right;margin-top:5px;margin-right:10px;"></div></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center">정렬번호</td>
	<td colspan="3"><input type="text" name="sortnum" value="<%=chkiif(sortnum="","0",sortnum)%>" size="2"/></td>
</tr>
<tr bgcolor="#FFFFFF" align="center">
    <td colspan="4"><input type="button" value=" 취 소 " onClick="jsgolist();"/><input type="button" value=" 저 장 " onClick="jsSubmit();"/></td>
</tr>
</form>
</table>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
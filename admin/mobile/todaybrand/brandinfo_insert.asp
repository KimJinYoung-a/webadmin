<%@ language=vbscript %>
<% option explicit %>
<%
'###############################################
' PageName : enjoy_insert.asp
' Discription : 모바일 enjoybanner_new
' History : 2014.06.23 이종화
'###############################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/classes/mobile/today_brandinfoCls.asp" -->
<%
Dim idx , mode
Dim srcSDT , srcEDT
Dim mainStartDate, mainEndDate
Dim sDt, sTm, eDt, eTm
Dim linkurl , ordertext
Dim stdt , eddt
Dim maincopy , subcopy , mainimg , moreimg , isusing
Dim itemid1 , itemid2 , iteminfo
Dim itemname1 ,  itemname2
Dim itemimg1 ,  itemimg2 , makerid
dim tmpArr

	idx = requestCheckvar(request("idx"),16)
	srcSDT = request("sDt")
	srcEDT = request("eDt")

If idx = "" Then
	mode = "add"
Else
	mode = "modify"
End If

'// 수정시
If idx <> "" then
	dim oBrandinfo
	set oBrandinfo = new CMainbanner
	oBrandinfo.FRectIdx = idx
	oBrandinfo.GetOneContentsNew()

	idx				=	oBrandinfo.FOneItem.Fidx
	mainStartDate	=	oBrandinfo.FOneItem.Fstartdate
	mainEndDate		=	oBrandinfo.FOneItem.Fenddate
	isusing			=	oBrandinfo.FOneItem.Fisusing
	ordertext		=	oBrandinfo.FOneItem.Fordertext
	makerid			=	oBrandinfo.FOneItem.Fmakerid
	linkurl			=	oBrandinfo.FOneItem.Flinkurl
	maincopy		=	oBrandinfo.FOneItem.Fmaincopy
	subcopy			=	oBrandinfo.FOneItem.Fsubcopy
	mainimg			=	oBrandinfo.FOneItem.Fmainimg
	moreimg			=	oBrandinfo.FOneItem.Fmoreimg
	itemid1			=	oBrandinfo.FOneItem.Fitemid1
	itemid2			=	oBrandinfo.FOneItem.Fitemid2
	iteminfo		=	oBrandinfo.FOneItem.Fiteminfo

	set oBrandinfo = Nothing

	''response.write "작업중<br><br>"
	''response.write iteminfo & "작업중"
	''response.end

	Dim ii
	If ubound(Split(iteminfo,",-,")) > 0 Then ' 이미지 3개 정보
		'// 상품명 : Minions 보조배터리 3,350mAh 3종
		tmpArr = Split(iteminfo,",-,")
		For ii = 0 To ubound(tmpArr)
			If CStr(itemid1) = CStr(Split(tmpArr(ii),"|")(0)) Then
				itemname1 = Split(tmpArr(ii),"|")(1)
				itemimg1 =  webImgUrl & "/image/small/" & GetImageSubFolderByItemid(itemid1) & "/" & Split(tmpArr(ii),"|")(2)
			End If

			If CStr(itemid2) = CStr(Split(tmpArr(ii),"|")(0)) Then
				itemname2 = Split(tmpArr(ii),"|")(1)
				itemimg2 =  webImgUrl & "/image/small/" & GetImageSubFolderByItemid(itemid2) & "/" & Split(tmpArr(ii),"|")(2)
			End If
		Next
	End If
End If

dim dateOption
dateOption = request("dateoption")

if Not(mainStartDate="" or isNull(mainStartDate)) then
	sDt = left(mainStartDate,10)
	sTm = Num2Str(hour(mainStartDate),2,"0","R") &":"& Num2Str(minute(mainStartDate),2,"0","R") &":"& Num2Str(second(mainStartDate),2,"0","R")
else
	if srcSDT<>"" then
		sDt = left(srcSDT,10)
	elseif dateOption <> "" then
		sDt = dateOption
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
	elseif dateOption <> "" then
		eDt = dateOption
	else	
		eDt = date
	end if
	eTm = "23:59:59"
end If

%>
<link href="/js/jqueryui/css/jquery-ui.css" rel="stylesheet">
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/jqueryui/jquery-ui-1.10.2.custom.min.js"></script>
<script type='text/javascript'>
	function jsSubmit(){
var frm = document.frm;

if (frm.makerid.value == ""){
	alert("브랜드ID를 넣어주세요.");
	frm.makerid.focus();
	return;
}
if (frm.maincopy.value == ""){
	alert("메인카피를 넣어주세요.");
	frm.maincopy.focus();
	return;
}
if (frm.subcopy.value == ""){
	alert("서브카피를 넣어주세요.");
	frm.subcopy.focus();
	return;
}
if (frm.itemid1.value == ""){
	alert("상품코드1를 넣어주세요.");
	frm.itemid1.focus();
	return;
}
if (frm.itemid2.value == ""){
	alert("상품코드2를 넣어주세요.");
	frm.itemid2.focus();
	return;
}

if (confirm('저장 하시겠습니까?')){
	//frm.target = "blank";
	frm.submit();
}
	}

	function jsgolist(){
	self.location.href="/admin/mobile/todaybrand/";
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

// 상품정보 접수
function fnGetItemInfo(iid,v) {
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
			var rst = "<img src='" + $(xml).find("itemInfo").find("item").find("smallImage").text() + "' height='70' /><br/>"
				rst += $(xml).find("itemInfo").find("item").find("itemname").text();
				$("#lyItemInfo"+v).fadeIn();
				$("#lyItemInfo"+v).html(rst);
			} else {
				$("#lyItemInfo"+v).fadeOut();
			}
		},
		error: function(xhr, status, error) {
			alert("상품코드를 다시 입력 해주세요");
			return;
			/*alert(xhr + '\n' + status + '\n' + error);*/
		}
	});
}

function putLinkText(key) {
	var frm = document.frm;
	switch(key) {
case 'event':
	frm.linkurl.value='/event/eventmain.asp?eventid=이벤트번호';
	break;
case 'itemid':
	frm.linkurl.value='/category/category_itemprd.asp?itemid=상품코드';
	break;
case 'category':
	frm.linkurl.value='/category/category_list.asp?disp=카테고리';
	break;
case 'brand':
	frm.linkurl.value='/street/street_brand.asp?makerid=브랜드아이디';
	break;
	}
}

//브랜드 ID 검색 팝업창
function jsSearchBrandIDNew(frmName,compName){
	var compVal = "";
	try{
		compVal = eval("document.all." + frmName + "." + compName).value;
	}catch(e){
		compVal = "";
	}

	var popwin = window.open("popBrandSearch.asp?frmName=" + frmName + "&compName=" + compName + "&rect=" + compVal,"popBrandSearch","width=800 height=400 scrollbars=yes resizable=yes");

	popwin.focus();
}

//브랜드 이미지 검색 팝업창
function jsSearchBrandImage(frmName){
	var popwin = window.open("/admin/brand/brandimage/image_list.asp?mode=img&frmName="+frmName,"popBrandimgSearch","width=800 height=400 scrollbars=yes resizable=yes");
	popwin.focus();
}

// 상품검색 등록
function addnewItem(target) {
	var popwin; 		
	popwin = window.open("item_regist.asp?formName=frm&target="+target, "popup_item", "width=1024,height=768,scrollbars=yes,resizable=yes");
	popwin.focus();
}
</script>
<table width="80%" cellpadding="2" cellspacing="1" class="a" bgcolor="#3d3d3d">
<form name="frm" method="post" action="dobrandinfo.asp" style="margin:0px;">
<input type="hidden" name="mode" value="<%=mode%>">
<input type="hidden" name="adminid" value="<%=session("ssBctId")%>">
<input type="hidden" name="idx" value="<%=idx%>">
<input type="hidden" name="menupos" value="<%=menupos%>">
<tr bgcolor="#FFFFFF">
    <td bgcolor="#FFF999" align="center" width="10%">노출기간</td>
    <td colspan="3">
		<input type="text" id="sDt" name="StartDate" size="10" value="<%=sDt%>" />
		<input type="text" name="sTm" size="8" value="<%=sTm%>" /> ~
		<input type="text" id="eDt" name="EndDate" size="10" value="<%=eDt%>" />
		<input type="text" name="eTm" size="8" value="<%=eTm%>" />
    </td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center" width="10%">브랜드ID</td>
    <td colspan="3">
		<% NewDrawSelectBoxDesignerwithNameEvent "makerid", makerid %>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center" width="10%">메인카피</td>
	<td colspan="3"><input type="text" name="maincopy" value="<%=maincopy%>" size="60"/></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center" width="10%">서브카피</td>
	<td colspan="3"><textarea name="subcopy" cols="80" rows="8"/><%=subcopy%></textarea></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center" width="10%">브랜드 배너</td>
	<td align="left">
		<input type="hidden" name="mainimg" value="<%=mainimg%>">
		<% If mainimg <> "" Then %>
		<br/><img src="<%=mainimg%>" width="200" id="mainimg" /><br>
		<% Else %>
		<br/><img src="/images/admin_login_logo2.png" width="200" border="0" id="mainimg"></br><span id="imgurl"></span><br>
		<% End If %>
		<input type="button" value="이미지 불러오기" onClick="jsSearchBrandImage(this.form.name);"/>
	</td>
	<td bgcolor="#FFF999" align="center" width="10%">더보기 배너</td>
	<td align="left">
		<input type="hidden" name="moreimg" id="moreimg" value="<%=moreimg%>">
		<% If moreimg <> "" Then %>
		<br/><img src="<%=moreimg%>" width="120" height="120" id="moreimgsrc" />
		<% Else %>
		<br/><img src="/images/admin_login_logo2.png" width="120" height="120" id="moreimgsrc" /></br>
		<% End If %>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
    <td bgcolor="#FFD99D" align="center">상품코드1</td>
    <td>
        <input type="text" name="itemid1" id="itemid1" value="<%=itemid1%>" size="8" maxlength="8" class="text" require="N" onClick="addnewItem('itemid1');" title="상품코드" />
		<div id="lyItemInfo1" style="display:<%=chkIIF(itemid1="","none","")%>;">
		<%
			if Not(itemName1="" or isNull(itemName1)) then
				Response.Write "<img src='" & itemimg1 & "' height='70' id='item1img' /><br/>"
				Response.Write itemName1
			end if
		%>
		</div>
    </td>
	<td bgcolor="#FFD99D" align="center">상품코드2</td>
    <td>
        <input type="text" name="itemid2" id="itemid2" value="<%=itemid2%>" size="8" maxlength="8" class="text" require="N" onClick="addnewItem('itemid2');"  title="상품코드" />
        <div id="lyItemInfo2" style="display:<%=chkIIF(itemid2="","none","")%>;">
		<%
			if Not(itemName2="" or isNull(itemName2)) then
				Response.Write "<img src='" & itemimg2 & "' height='70' id='item1img'/><br/>"
				Response.Write itemName2
			end if
		%>
		</div>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center">더보기 링크 URL</td>
	<td align="left" colspan="3">
	<input type="text" name="linkurl" value="<%=linkurl%>" maxlength="128" style="width:100%">
	<font color="#707070">
	- <font color="red"><strong>app & mobile 공용</strong></font> - <br/>
	- <span style="cursor:pointer" onClick="putLinkText('event');">이벤트 링크 : /event/eventmain.asp?eventid=<font color="darkred">이벤트코드</font></span><br/>
	- <span style="cursor:pointer" onClick="putLinkText('itemid');">상품코드 링크 : /category/category_itemprd.asp?itemid=<font color="darkred">상품코드 (O)</font></span><br/>
	- <span style="cursor:pointer" onClick="putLinkText('category');">카테고리 링크 : /category/category_list.asp?disp=<font color="darkred">카테고리</font></span><br/>
	- <span style="cursor:pointer" onClick="putLinkText('brand');">브랜드아이디 링크 : /street/street_brand.asp?makerid=<font color="darkred">브랜드아이디</font></span><br/>
	</font>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center">사용여부</td>
	<td colspan="3"><div style="float:left;"><input type="radio" name="isusing" value="Y" <%=chkiif(isusing = "Y","checked","")%> checked />사용함 &nbsp;&nbsp;&nbsp; <input type="radio" name="isusing" value="N"  <%=chkiif(isusing = "N","checked","")%>/>사용안함</div> <div style="float:right;margin-top:5px;margin-right:10px;"></div></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center">작업자 지시사항</td>
	<td colspan="3"><textarea name="ordertext" cols="80" rows="8"/><%=ordertext%></textarea></td>
</tr>
<tr bgcolor="#FFFFFF" align="center">
    <td colspan="4"><input type="button" value=" 취 소 " onClick="jsgolist();"/><input type="button" value=" 저 장 " onClick="jsSubmit();"/></td>
</tr>
</form>
</table>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->

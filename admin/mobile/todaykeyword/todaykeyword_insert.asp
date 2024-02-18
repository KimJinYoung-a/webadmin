<%@ language=vbscript %>
<% option explicit %>
<%
'###############################################
' PageName : todaykeyword_insert.asp
' Discription : 모바일 투데이 키워드
' History : 2017-08-04 이종화 생성
'###############################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/classes/mobile/today_keywordCls.asp" -->
<%
Dim idx , mode
Dim srcSDT , srcEDT
Dim mainStartDate, mainEndDate
Dim sDt, sTm, eDt, eTm
Dim linkurl , ordertext
Dim stdt , eddt , isusing , ver_no , keyword , picknum

Dim itemid1 , itemid2 , itemid3 , itemid4 , iteminfo
Dim itemname1 ,  itemname2 , itemName3 , itemName4
Dim itemimg1 ,  itemimg2 , itemimg3 , itemimg4
Dim bgcolor

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
	dim oKeyword
	set oKeyword = new CMainbanner
	oKeyword.FRectIdx = idx
	oKeyword.GetOneContents()

	idx				=	oKeyword.FOneItem.Fidx
	mainStartDate	=	oKeyword.FOneItem.Fstartdate
	mainEndDate		=	oKeyword.FOneItem.Fenddate
	isusing			=	oKeyword.FOneItem.Fisusing
	ordertext		=	oKeyword.FOneItem.Fordertext
	linkurl			=	oKeyword.FOneItem.Flinkurl
	ver_no			=	oKeyword.FOneItem.Fver_no
	keyword			=	oKeyword.FOneItem.Fkeyword
	picknum			=	oKeyword.FOneItem.Fpicknum
	itemid1			=	oKeyword.FOneItem.Fitemid1
	itemid2			=	oKeyword.FOneItem.Fitemid2
	itemid3			=	oKeyword.FOneItem.Fitemid3
	itemid4			=	oKeyword.FOneItem.Fitemid4
	itemimg1		=	oKeyword.FOneItem.Fitemimg1
	itemimg2		=	oKeyword.FOneItem.Fitemimg2
	itemimg3		=	oKeyword.FOneItem.Fitemimg3
	itemimg4		=	oKeyword.FOneItem.Fitemimg4
	iteminfo		=	oKeyword.FOneItem.Fiteminfo
	bgcolor			=	oKeyword.FOneItem.Fbgcolor

	set oKeyword = Nothing

	Dim ii
	If iteminfo <> "" and not isnull(iteminfo) Then
		If ubound(Split(iteminfo,"^^")) > 0 Then ' 이미지 3개 정보
			For ii = 0 To ubound(Split(iteminfo,","))
				If CStr(itemid1) = CStr(Split(Split(iteminfo,",")(ii),"|")(0)) And itemimg1 = "" Then
					itemname1 = Split(Split(iteminfo,",")(ii),"|")(1)
					itemimg1 =  webImgUrl & "/image/small/" & GetImageSubFolderByItemid(itemid1) & "/" & Split(Split(iteminfo,",")(ii),"|")(2)
				End If

				If CStr(itemid2) = CStr(Split(Split(iteminfo,",")(ii),"|")(0)) And itemimg2 = "" Then
					itemname2 = Split(Split(iteminfo,",")(ii),"|")(1)
					itemimg2 =  webImgUrl & "/image/small/" & GetImageSubFolderByItemid(itemid2) & "/" & Split(Split(iteminfo,",")(ii),"|")(2)
				End If

				If CStr(itemid3) = CStr(Split(Split(iteminfo,",")(ii),"|")(0)) And itemimg3 = "" Then
					itemname3 = Split(Split(iteminfo,",")(ii),"|")(1)
					itemimg3 =  webImgUrl & "/image/small/" & GetImageSubFolderByItemid(itemid3) & "/" & Split(Split(iteminfo,",")(ii),"|")(2)
				End If

				If CStr(itemid4) = CStr(Split(Split(iteminfo,",")(ii),"|")(0)) And itemimg4 = "" Then
					itemname4 = Split(Split(iteminfo,",")(ii),"|")(1)
					itemimg4 =  webImgUrl & "/image/small/" & GetImageSubFolderByItemid(itemid4) & "/" & Split(Split(iteminfo,",")(ii),"|")(2)
				End If
			Next
		End If
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
	elseif dateoption <> "" then
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
	elseif dateoption <> "" then
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

		if (frm.ver_no.value == ""){
			alert("Ver.을 넣어주세요.");
			frm.ver_no.focus();
			return;
		}
		if (frm.keyword.value == ""){
			alert("키워드를 넣어주세요.");
			frm.keyword.focus();
			return;
		}

		if (frm.linkurl.value == "" || frm.linkurl.value == "/search/search_item.asp?rect=키워드" || frm.linkurl.value == "/event/eventmain.asp?eventid=이벤트번호" || frm.linkurl.value == "/category/category_itemprd.asp?itemid=상품코드" || frm.linkurl.value == "/category/category_list.asp?disp=카테고리" || frm.linkurl.value == "/street/street_brand.asp?makerid=브랜드아이디" ){
			alert("링크 URL을 넣어주세요.");
			frm.linkurl.focus();
			return;
		}
		if (frm.itemid1.value == ""){
			alert("상품코드1를 넣어주세요.");
			return;
		}

		if (frm.itemid2.value == ""){
			alert("상품코드2를 넣어주세요.");
			return;
		}

		if (frm.itemid3.value == ""){
			alert("상품코드3를 넣어주세요.");
			return;
		}

		if (frm.itemid4.value == ""){
			alert("상품코드4를 넣어주세요.");
			return;
		}

		if (confirm('저장 하시겠습니까?')){
			//frm.target = "blank";
			frm.submit();
		}
	}

	function jsgolist(){
	self.location.href="/admin/mobile/todaykeyword/";
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
			var rst = "<img src='" + $(xml).find("itemInfo").find("item").find("smallImage").text() + "' height='100' /><br/>"
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
		case 'search':
			frm.linkurl.value='/search/search_item.asp?rect=키워드';
			break;
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

//링크 URL
<% if mode = "add" then %>
$(function(){
	var frmtext = $("#keywordtext");
	frmtext.bind("keyup",function(){
		if ($(this).val().length > 0){
			$("#linkurl").val("/search/search_item.asp?rect=" + $(this).val());
		}else{
			$("#linkurl").val("");
		}
	});
});
<% end if %>
</script>
<table width="80%" cellpadding="2" cellspacing="1" class="a" bgcolor="#3d3d3d">
<form name="frm" method="post" action="<%=uploadUrl%>/linkweb/mobile/todaykeyword_proc.asp" enctype="multipart/form-data" style="margin:0px;">
<input type="hidden" name="mode" value="<%=mode%>">
<input type="hidden" name="adminid" value="<%=session("ssBctId")%>">
<input type="hidden" name="idx" value="<%=idx%>">
<input type="hidden" name="menupos" value="<%=menupos%>">
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center" width="10%">Ver.</td>
    <td colspan="3">
		<input type="text" name="ver_no" size="8" value="<%=ver_no%>" />
    </td>
</tr>
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
	<td bgcolor="#FFF999" align="center" width="10%">키워드(타이틀)</td>
    <td colspan="3">
		# <input type="text" name="keyword" size="20" value="<%=keyword%>" id="keywordtext" maxlength="5"/><font color="red"><strong>※ 최대 5자 제한 ※</strong></font>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center">링크 URL</td>
	<td align="left" colspan="3">
	<input type="text" name="linkurl" value="<%=linkurl%>" maxlength="128" style="width:50%" id="linkurl"><br/>
	<font color="#707070">
	- <font color="red"><strong>app & mobile 공용</strong></font> - <br/>
	- <span style="cursor:pointer" onClick="putLinkText('search');">검색 링크 : /search/search_item.asp?rect=<font color="darkred">키워드</font></span><br/>
	- <span style="cursor:pointer" onClick="putLinkText('event');">이벤트 링크 : /event/eventmain.asp?eventid=<font color="darkred">이벤트코드</font></span><br/>
	- <span style="cursor:pointer" onClick="putLinkText('itemid');">상품코드 링크 : /category/category_itemprd.asp?itemid=<font color="darkred">상품코드 (O)</font></span><br/>
	- <span style="cursor:pointer" onClick="putLinkText('category');">카테고리 링크 : /category/category_list.asp?disp=<font color="darkred">카테고리</font></span><br/>
	- <span style="cursor:pointer" onClick="putLinkText('brand');">브랜드아이디 링크 : /street/street_brand.asp?makerid=<font color="darkred">브랜드아이디</font></span><br/>
	</font>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center">작업자 지시사항</td>
	<td colspan="3"><textarea name="ordertext" cols="80" rows="8"/><%=ordertext%></textarea></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center">사용여부</td>
	<td colspan="3"><div style="float:left;"><input type="radio" name="isusing" value="Y" <%=chkiif(isusing = "Y","checked","")%> checked />사용함 &nbsp;&nbsp;&nbsp; <input type="radio" name="isusing" value="N"  <%=chkiif(isusing = "N","checked","")%>/>사용안함</div> <div style="float:right;margin-top:5px;margin-right:10px;"></div></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center" colspan="4">등록하는 배너가 우선 노출 됩니다(이미지 없을 경우 상품코드 이미지).</td>
</tr>
<tr bgcolor="#FFFFFF">
    <td bgcolor="#FFD99D" align="center">상품코드1</td>
    <td>
        <input type="text" name="itemid1" value="<%=itemid1%>" size="8" maxlength="8" class="text" require="N" onblur="fnGetItemInfo(this.value,'1')" title="상품코드" />
    </td>
	<td bgcolor="#FFD99D" align="center">상품코드2</td>
    <td>
        <input type="text" name="itemid2" value="<%=itemid2%>" size="8" maxlength="8" class="text" require="N" onblur="fnGetItemInfo(this.value,'2')" title="상품코드" />
    </td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center" width="10%">상품이미지1</td>
	<td align="left">
		<input type="file" name="itemimg1" class="file" require="N" style="width:80%;" /> <input type="checkbox" name="delimg1" value="Y"/> : 삭제
		<div id="lyItemInfo1" style="display:<%=chkIIF(itemid1="","none","")%>;">
		<%
			if Not(itemimg1="" or isNull(itemimg1)) then
				Response.Write "<img src='" & itemimg1 & "' height='100' /><br/>"
				Response.Write itemName1
			end if
		%>
		</div>
		<br/><%=itemimg1%>
	</td>
	<td bgcolor="#FFF999" align="center" width="10%">상품이미지2</td>
	<td align="left">
		<input type="file" name="itemimg2" class="file" require="N" style="width:80%;" /> <input type="checkbox" name="delimg2" value="Y"/> : 삭제
		<div id="lyItemInfo2" style="display:<%=chkIIF(itemid2="","none","")%>;">
		<%
			if Not(itemimg2="" or isNull(itemimg2)) then
				Response.Write "<img src='" & itemimg2 & "' height='100' /><br/>"
				Response.Write itemName2
			end if
		%>
		</div>
		<br/><%=itemimg2%>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
    <td bgcolor="#FFD99D" align="center">상품코드3</td>
    <td>
        <input type="text" name="itemid3" value="<%=itemid3%>" size="8" maxlength="8" class="text" require="N" onblur="fnGetItemInfo(this.value,'3')" title="상품코드" />
    </td>
	<td bgcolor="#FFD99D" align="center">상품코드4</td>
    <td>
        <input type="text" name="itemid4" value="<%=itemid4%>" size="8" maxlength="8" class="text" require="N" onblur="fnGetItemInfo(this.value,'4')" title="상품코드" />
    </td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center" width="10%">상품이미지3</td>
	<td align="left">
		<input type="file" name="itemimg3" class="file" require="N" style="width:80%;" /> <input type="checkbox" name="delimg3" value="Y"/> : 삭제
		<div id="lyItemInfo3" style="display:<%=chkIIF(itemid3="","none","")%>;">
		<%
			if Not(itemimg3="" or isNull(itemimg3)) then
				Response.Write "<img src='" & itemimg3 & "' height='100' /><br/>"
				Response.Write itemName3
			end if
		%>
		</div>
		<br/><%=itemimg3%>
	</td>
	<td bgcolor="#FFF999" align="center" width="10%">상품이미지4</td>
	<td align="left">
		<input type="file" name="itemimg4" class="file" require="N" style="width:80%;" /> <input type="checkbox" name="delimg4" value="Y"/> : 삭제
		<div id="lyItemInfo4" style="display:<%=chkIIF(itemid4="","none","")%>;">
		<%
			if Not(itemimg4="" or isNull(itemimg4)) then
				Response.Write "<img src='" & itemimg4 & "' height='100' /><br/>"
				Response.Write itemName4
			end if
		%>
		</div>
		<br/><%=itemimg4%>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center">PICK 사용 여부</td>
	<td>
		<input type="radio" id="pick1" name="picknum" value="0" checked <%=chkiif(picknum=0,"checked","")%>/> <label for="pick1">사용안함</label>
		&nbsp <input type="radio" id="pick2" name="picknum" value="1" <%=chkiif(picknum=1,"checked","")%>/> <label for="pick2">배너1</label>
		&nbsp <input type="radio" id="pick3" name="picknum" value="2" <%=chkiif(picknum=2,"checked","")%>/> <label for="pick3">배너2</label>
		&nbsp <input type="radio" id="pick4" name="picknum" value="3" <%=chkiif(picknum=3,"checked","")%>/> <label for="pick4">배너3</label>
		&nbsp <input type="radio" id="pick5" name="picknum" value="4" <%=chkiif(picknum=4,"checked","")%>/> <label for="pick5">배너4</label>
	</td>
	<td bgcolor="#FFF999" align="center">배경색</td>
	<td>
		<input type="text" name="bgcolor" value="<%=bgcolor%>"/>#붙여주세요 ex)#000000
	</td>
</tr>
<tr bgcolor="#FFFFFF" align="center">
    <td colspan="4"><input type="button" value=" 취 소 " onClick="jsgolist();"/><input type="button" value=" 저 장 " onClick="jsSubmit();"/></td>
</tr>
</form>
</table>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
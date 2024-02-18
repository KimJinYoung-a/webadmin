<%@ language=vbscript %>
<% option explicit %>
<%
'###############################################
' PageName : deal_insert.asp
' Discription : 모바일 dealbanner_new
' History : 2014.06.23 이종화
'###############################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/classes/mobile/todaydealCls.asp" -->
<%
'###############################################
'이벤트 신규 등록시
'###############################################
%>
<!-- #include virtual="/admin/eventmanage/common/event_function.asp"-->
<!-- #include virtual="/lib/classes/event/eventmanageCls.asp"-->
<%
Dim eCode
Dim idx , isusing , mode
Dim srcSDT , srcEDT
Dim mainStartDate, mainEndDate
Dim sDt, sTm, eDt, eTm
Dim itemurl , itemurlmo
Dim dealtitle
Dim prevDate
Dim itemid , itemname , limitno
Dim stdt , eddt , sortnum , smallImage
Dim gubun1 , gubun2 , limityn

	idx = requestCheckvar(request("idx"),16)
	srcSDT = request("sDt")
	srcEDT = request("eDt")
	prevDate = request("prevDate")

If idx = "" Then
	mode = "add"
Else
	mode = "modify"
End If

'// 수정시
If idx <> "" then
	dim oTodayDealOne
	set oTodayDealOne = new CMainbanner
	oTodayDealOne.FRectIdx = idx
	oTodayDealOne.GetOneContents()

	idx					=	oTodayDealOne.FOneItem.Fidx
	smallImage			=	oTodayDealOne.FOneItem.FSmallimg
	itemurl				=	oTodayDealOne.FOneItem.Fitemurl
	itemurlmo			=	oTodayDealOne.FOneItem.Fitemurlmo '2014-09-16 모바일용추가
	dealtitle			=	oTodayDealOne.FOneItem.Fdealtitle
	mainStartDate		=	oTodayDealOne.FOneItem.Fstartdate
	mainEndDate			=	oTodayDealOne.FOneItem.Fenddate
	isusing				=	oTodayDealOne.FOneItem.Fisusing
	sortnum				=	oTodayDealOne.FOneItem.Fsortnum
	gubun1				=	oTodayDealOne.FOneItem.Fgubun1
	gubun2				=	oTodayDealOne.FOneItem.Fgubun2
	limityn				=	oTodayDealOne.FOneItem.Flimityn
	limitno				=	oTodayDealOne.FOneItem.Flimitno
	itemid				=	oTodayDealOne.FOneItem.Fitemid
	itemname			=	oTodayDealOne.FOneItem.Fitemname


	set oTodayDealOne = Nothing
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
end If

%>
<link href="/js/jqueryui/css/jquery-ui.css" rel="stylesheet">
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/jqueryui/jquery-ui-1.10.2.custom.min.js"></script>
<script type='text/javascript'>

	function getByteLength(str) {
		var p, len = 0;
		for (p = 0; p < str.length; p++) {
			(str.charCodeAt(p)  > 255) ? len += 2 : len++;
		}

		return len;
	}

	function jsSubmit(){
		var frm = document.frm;

		if (getByteLength(frm.dealtitle.value) >= 50) {
			alert("제목을 짧게 입력하세요(" + getByteLength(frm.dealtitle.value) + "/50)");
			return;
		}

		if (confirm('저장 하시겠습니까?')){
			//frm.target = "blank";
			frm.submit();
		}
	}
	function jsgolist(){
		self.location.href="/admin/appmanage/today/todaydeal/";
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

	function onchgbox(v){
		if (v == "3"){
			$("#gubun2").css("display","block");
		}else{
			$("#gubun2").css("display","none");
		}
	}

	function fnGetItemInfo(iid) {
		$.ajax({
			type: "GET",
			url: "act_iteminfo.asp?itemid="+iid,
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
					var rst = "<img src='" + $(xml).find("itemInfo").find("item").find("smallImage").text() + "' height='80' /><br/><br/>한정수량 : " + $(xml).find("itemInfo").find("item").find("limitno").text() + "개 <br/>"
						rst += "상품명 : <input type='text' value='" + $(xml).find("itemInfo").find("item").find("itemname").text() + "' size='40' id='tempname' onkeyup='copytxt();'/>"
					$("#lyItemInfo").fadeIn();
					$("#lyItemInfo").html(rst);
					$("#itemname").val($(xml).find("itemInfo").find("item").find("itemname").text());
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
	//아이템이름복사
	function copytxt(){
		var txt1 = $("#tempname");
		var txt2 = $("#itemname");
		txt2.val(txt1.val());
	}
</script>
<table width="1000" cellpadding="2" cellspacing="1" class="a" bgcolor="#3d3d3d">
<form name="frm" method="post" action="dotodaydeal.asp">
<input type="hidden" name="mode" value="<%=mode%>">
<input type="hidden" name="adminid" value="<%=session("ssBctId")%>">
<input type="hidden" name="idx" value="<%=idx%>">
<input type="hidden" name="prevDate" value="<%=prevDate%>">
<input type="hidden" name="menupos" value="<%=menupos%>">
<input type="hidden" id="itemname" name="itemname" value="<%=itemname%>">
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
	<td bgcolor="#FFF999" align="center" height="40">구분명</td>
	<td colspan="3">
		<div style="float:left">
		<select name="gubun1" onchange="onchgbox(this.value);" width="100">
			<option value="">=====구분선택=====</option>
			<option value="1" <%=chkiif(gubun1="1","selected","")%>>TIME SALE</option>
			<option value="2" <%=chkiif(gubun1="2","selected","")%>>WISH NO.1</option>
			<option value="3" <%=chkiif(gubun1="3","selected","")%>>ISSUE ITEM</option>
		</select>&nbsp;&nbsp;
		</div>
		<div>
		<select id="gubun2" name="gubun2" style="display:<%=chkiif(gubun1="3","block","none")%>;" width="100">
			<option value="">=====이슈선택=====</option>
			<option value="1" <%=chkiif(gubun2="1","selected","")%>>한정 재입고</option>
			<option value="2" <%=chkiif(gubun2="2","selected","")%>>HOT ITEM</option>
			<option value="3" <%=chkiif(gubun2="3","selected","")%>>SPECIAL EDITION</option>
			<option value="4" <%=chkiif(gubun2="4","selected","")%>>10x10 ONLY</option>
		</select>
		</div>
		<div style="clear:both;"></div>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center" height="40">제목</td>
	<td colspan="3">
		<input type="text" name="dealtitle" size="50" value="<%=dealtitle%>"/>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center" height="40">상품</td>
	<td colspan="3">
		<input type="text" name="itemid" value="<%= itemid %>" size="8" maxlength="8" class="text" require="N" onblur="fnGetItemInfo(this.value)" title="상품코드" />
		<input type="button" value="상품등록"  >
        <div id="lyItemInfo" style="display:<%=chkIIF(itemid="","none","block")%>;">
        <%
        	if Not(itemName="" or isNull(itemName)) then
        		Response.Write "<img src='" & smallImage & "' height='80' /><br/><br/>한정수량 :" & limitno & "개<br/>"
	    		Response.Write "상품명 : <input type='text' value='"& itemName &"' id='tempname' onkeyup='copytxt();' size='40' />"
        	end if
        %>
        </div>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center" height="40">모바일상품URL</td>
	<td colspan="3">
		<input type="text" name="itemurlmo" size="110" value="<%=itemurlmo%>"/><br/><br/>
		<span style="color:red">※ 쿠폰상품이 아닌 일반 상품의 경우도 URL을 넣어 주세요<br/>
		예) http://m.10x10.co.kr/category/category_itemprd.asp?itemid=884073</span>
		<br/><br/>
		※ [ON]상품관리 &gt;&gt; 상품쿠폰관리 대상 상품 팝업 클릭 &gt;&gt; 상품번호 하단 URL생성 &gt;&gt; 모바일웹 링크 참고<br/>
		예) http://m.10x10.co.kr/category/category_itemprd.asp?itemid=663507&ldv=MzIwMCAg
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center" height="40">APP상품URL</td>
	<td colspan="3">
		<input type="text" name="itemurl" size="110" value="<%=itemurl%>"/><br/><br/>
		<span style="color:red">※ 쿠폰상품이 아닌 일반 상품의 경우도 URL을 넣어 주세요<br/>
		예) http://m.10x10.co.kr/apps/appcom/wish/web2014/category/category_itemprd.asp?itemid=884073</span>
		<br/><br/>
		※ [ON]상품관리 &gt;&gt; 상품쿠폰관리 대상 상품 팝업 클릭 &gt;&gt; 상품번호 하단 URL생성 &gt;&gt; wishApp 링크 참고<br/>
		예) http://m.10x10.co.kr/apps/appcom/wish/web2014/category/category_itemprd.asp?itemid=884073&ldv=OTQ4MiAg
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center">정렬 번호</td>
	<td colspan="3"><input type="text" name="sortnum" size="10" value="99" maxlength="3"/></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center">사용여부</td>
	<td colspan="3"><div style="float:left;"><input type="radio" name="isusing" value="Y" <%=chkiif(isusing = "Y","checked","")%> checked />사용함 &nbsp;&nbsp;&nbsp; <input type="radio" name="isusing" value="N"  <%=chkiif(isusing = "N","checked","")%>/>사용안함</div> <div style="float:right;margin-top:5px;margin-right:10px;"></div></td>
</tr>
<tr bgcolor="#FFFFFF" align="center">
    <td colspan="4"><input type="button" value=" 취 소 " onClick="jsgolist();"/><input type="button" value=" 저 장 " onClick="jsSubmit();"/></td>
</tr>
</form>
</table>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->

<%@ language=vbscript %>
<% option explicit %>
<%
'###############################################
' PageName : todaytheme_insert.asp
' Discription : 모바일 todaytheme
' History : 2014.01.07 이종화
'###############################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/classes/mobile/todaythemeCls.asp" -->
<%
Dim idx , subImage1 , subImage2 , subImage3 , isusing , mode
Dim srcSDT , srcEDT 
Dim mainStartDate, mainEndDate , lp
Dim sDt, sTm, eDt, eTm , gubun , maintitle , subtitle , prevDate
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
	dim oTodayThemeList
	set oTodayThemeList = new Ctodaytheme
	oTodayThemeList.FRectIdx = idx
	oTodayThemeList.GetOneContents()

	gubun	=	oTodayThemeList.FOneItem.Fgubun
	maintitle	=	oTodayThemeList.FOneItem.Fmaintitle
	mainStartDate	=	oTodayThemeList.FOneItem.Fstartdate
	mainEndDate	=	oTodayThemeList.FOneItem.Fenddate 
	isusing			=	oTodayThemeList.FOneItem.Fisusing
	subtitle			=	oTodayThemeList.FOneItem.Fsubtitle

	set oTodayThemeList = Nothing
End If 

Dim oSubItemList
set oSubItemList = new Ctodaytheme
	oSubItemList.FPageSize = 100
	oSubItemList.FRectlistIdx = idx
	If idx <> "" then
		oSubItemList.GetContentsItemList()
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
	
		if (confirm('저장 하시겠습니까?')){
			//frm.target = "blank";
			frm.submit();
		}
	}
	function jsgolist(){
		self.location.href="/admin/mobile/todaytheme/";
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
	
	//라디오버튼
    $(".rdoUsing").buttonset().children().next().attr("style","font-size:11px;");

	$( "#subList" ).sortable({
		placeholder: "ui-state-highlight",
		start: function(event, ui) {
			ui.placeholder.html('<td height="54" colspan="10" style="border:1px solid #F9BD01;">&nbsp;</td>');
		},
		stop: function(){
			var i=99999;
			$(this).parent().find("input[name^='sort']").each(function(){
				if(i>$(this).val()) i=$(this).val()
			});
			if(i<=0) i=1;
			$(this).parent().find("input[name^='sort']").each(function(){
				$(this).val(i);
				i++;
			});
		}
	});

});

function nextdate(date){
	var ndate = date
	var dateinfo = ndate.split("-");
	var tdate =  new Date(dateinfo[0],dateinfo[1]-1,dateinfo[2]);

	var tyear = tdate.getFullYear(dateinfo[0]);
	var tmonth=tdate.getMonth(dateinfo[1]-1) + 1;
	tmonth = (tmonth<10)? '0'+ tmonth : tmonth;
	var tday=tdate.getDate(dateinfo[2])+1; 
	tday = (tday<10)? '0'+ tday : tday;

	var nextdate = tyear+'-'+tmonth+'-'+tday;

	frm.eDt.value = nextdate;
}
//소재
function popSubEdit(subidx) {
<% if idx <>"" then %>
    var popwin = window.open('popSubItemEdit.asp?listIdx=<%=idx%>&subIdx='+subidx,'popTemplateManage','width=800,height=300,scrollbars=yes,resizable=yes');
    popwin.focus();
<% else %>
	alert("템플릿 컨텐츠 정보를 먼저 등록해주세요.");
<% end if %>
}

// 상품검색 일괄 등록
function popRegSearchItem() {
<% if idx <> "" then %>
    var popwin = window.open("/admin/itemmaster/pop_itemAddInfo.asp?sellyn=Y&usingyn=Y&defaultmargin=0&acURL=/admin/mobile/todaytheme/doSubRegItemCdArray.asp?listidx=<%=idx%>", "popup_item", "width=800,height=500,scrollbars=yes,resizable=yes");
    popwin.focus();
<% else %>
	alert("템플릿 컨텐츠 정보를 먼저 등록해주세요.");
<% end if %>
}

// 상품코드 일괄 등록
function popRegArrayItem() {
<% if idx<>"" then %>
    var popwin = window.open('popSubRegItemCdArray.asp?listIdx=<%=idx%>','popRegArray','width=600,height=300,scrollbars=yes,resizable=yes');
    popwin.focus();
<% else %>
	alert("템플릿 컨텐츠 정보를 먼저 등록해주세요.");
<% end if %>
}

function chkAllItem() {
	if($("input[name='chkIdx']:first").attr("checked")=="checked") {
		$("input[name='chkIdx']").attr("checked",false);
	} else {
		$("input[name='chkIdx']").attr("checked","checked");
	}
}

function saveList() {
	var chk=0;
	$("form[name='frmList']").find("input[name='chkIdx']").each(function(){
		if($(this).attr("checked")) chk++;
	});
	if(chk==0) {
		alert("수정하실 소재를 선택해주세요.");
		return;
	}
	if(confirm("지정하신 목록의 선택 정보를 저장하시겠습니까?")) {
		document.frmList.action="doListModify.asp";
		document.frmList.submit();
	}
}
</script>
<table width="900" cellpadding="2" cellspacing="1" class="a" bgcolor="#3d3d3d">
<form name="frm" method="post" action="dotodaytheme.asp">
<input type="hidden" name="mode" value="<%=mode%>">
<input type="hidden" name="adminid" value="<%=session("ssBctId")%>">
<input type="hidden" name="idx" value="<%=idx%>">
<input type="hidden" name="prevDate" value="<%=prevDate%>">
<tr bgcolor="#FFFFFF">
	<% If idx = ""  Then %>
	<td colspan="2" align="center" height="35">등록 진행 중 입니다.</td>
	<% Else %>
	<td bgcolor="#FFF999" colspan="2" align="center" height="35">수정 진행 중 입니다.</td>
	<% End If %>
</tr>
<% If idx <> ""  Then %>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center" width="15%">idx</td>
	<td><%=idx%></td>
</tr>
<% End If %>
<tr bgcolor="#FFFFFF">
    <td bgcolor="#FFF999" align="center" width="15%">노출기간</td>
    <td>
		<input type="text" id="sDt" name="StartDate" size="10" value="<%=sDt%>" />
		<input type="text" name="sTm" size="8" value="<%=sTm%>" /> ~
		<input type="text" id="eDt" name="EndDate" size="10" value="<%=eDt%>" />
		<input type="text" name="eTm" size="8" value="<%=eTm%>" />
    </td>
</tr>

<tr bgcolor="#FFFFFF">
    <td bgcolor="#FFF999" align="center" width="15%">제목</td>
    <td>
		<input type="text" name="maintitle" size="100" value="<%If idx="" then%>수정요망--&gt; 오늘의 테마<% Else %><%=maintitle%><% End if%>" />
    </td>
</tr>

<tr bgcolor="#FFFFFF">
    <td bgcolor="#FFF999" align="center" width="15%">부제</td>
    <td>
		<input type="text" name="subtitle" size="100" value="<%=subtitle%>" />
    </td>
</tr>

<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center">사용여부</td>
	<td><div style="float:left;"><input type="radio" name="isusing" value="Y" <%=chkiif(isusing = "Y","checked","")%> checked />사용함 &nbsp;&nbsp;&nbsp; <input type="radio" name="isusing" value="N"  <%=chkiif(isusing = "N","checked","")%>/>사용안함</div> <div style="float:right;margin-top:5px;margin-right:10px;"></div></td>
</tr>

<tr bgcolor="#FFFFFF" align="center">
    <td colspan="4"><input type="button" value=" 취 소 " onClick="jsgolist();"/><input type="button" value=" 저 장 " onClick="jsSubmit();"/></td>
</tr>
</form>
</table>

<%
	If idx <> "" then
%>
<p><b>▶ 소재 정보</b></p>
<!-- // 등록된 소재 목록 --------->
<form name="frmList" method="POST" action="" style="margin:0;">
<input type="hidden" name="mode" value="sub">
<table width="900" align="left" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="#FFFFFF">
	<td colspan="10">
		<table width="100%" border="0" class="a" cellpadding="0" cellspacing="0">
		<tr>
		    <td align="left">
		    	총 <%=oSubItemList.FTotalCount%> 건 /
		    	<input type="button" value="전체선택" class="button" onClick="chkAllItem()">
		    	<input type="button" value="상태저장" class="button" onClick="saveList()" title="표시순서 및 사용여부를 일괄저장합니다.">
		    </td>
		    <td align="right">
		    	<input type="button" value="상품코드로 등록" class="button" onClick="popRegArrayItem()" />
		    	<input type="button" value="상품 추가" class="button" onClick="popRegSearchItem()" />
		    	<img src="/images/icon_new_registration.gif" border="0" onclick="popSubEdit('')" style="cursor:pointer;" align="absmiddle">
		    </td>
		</tr>
		</table>
	</td>
</tr>
<col width="30" />
<col width="60" />
<col span="3" width="0*" />
<col width="110" />
<col width="80" />
<col width="80" />
<tr align="center" bgcolor="#DDDDFF">
    <td>&nbsp;</td>
    <td>소재번호</td>
    <td>이미지</td>
    <td>상품코드</td>
    <td>표시순서</td>
    <td>사용여부</td>
</tr>
<tbody id="subList">
<%	For lp=0 to oSubItemList.FResultCount-1 %>
<tr align="center" bgcolor="<%=chkIIF(oSubItemList.FItemList(lp).FIsUsing="Y","#FFFFFF","#F3F3F3")%>#FFFFFF">
    <td><input type="checkbox" name="chkIdx" value="<%=oSubItemList.FItemList(lp).FsubIdx%>" /></td>
    <td onclick="popSubEdit(<%=oSubItemList.FItemList(lp).FsubIdx%>)" style="cursor:pointer;"><%=oSubItemList.FItemList(lp).FsubIdx%></td>
    <td onclick="popSubEdit(<%=oSubItemList.FItemList(lp).FsubIdx%>)" style="cursor:pointer;">
    <%
    	if Not(oSubItemList.FItemList(lp).FsmallImage="" or isNull(oSubItemList.FItemList(lp).FsmallImage)) then
    		Response.Write "<img src='" & oSubItemList.FItemList(lp).FsmallImage & "' height='50' />"
    	end if
    %>
    </td>
    <td>
    <%
    	if Not(oSubItemList.FItemList(lp).FItemid="0" or isNull(oSubItemList.FItemList(lp).FItemid) or oSubItemList.FItemList(lp).FItemid="") then
    		Response.Write "[" & oSubItemList.FItemList(lp).FItemid & "]" & oSubItemList.FItemList(lp).Fitemname
    	end if
    %>
    </td>
    <td><input type="text" name="sort<%=oSubItemList.FItemList(lp).FsubIdx%>" size="3" class="text" value="<%=oSubItemList.FItemList(lp).Fsortnum%>" style="text-align:center;" /></td>
    <td>
		<span class="rdoUsing">
		<input type="radio" name="use<%=oSubItemList.FItemList(lp).FsubIdx%>" id="rdoUsing<%=lp%>_1" value="Y" <%=chkIIF(oSubItemList.FItemList(lp).FIsUsing="Y","checked","")%> /><label for="rdoUsing<%=lp%>_1">사용</label><input type="radio" name="use<%=oSubItemList.FItemList(lp).FsubIdx%>" id="rdoUsing<%=lp%>_2" value="N" <%=chkIIF(oSubItemList.FItemList(lp).FIsUsing="N","checked","")%> /><label for="rdoUsing<%=lp%>_2">삭제</label>
		</span>
    </td>
</tr>
<% Next %>
</tbody>
</table>
</form>
<%
	End If 
	set oSubItemList = Nothing
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
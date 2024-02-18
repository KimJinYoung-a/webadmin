<%@ language=vbscript %>
<% option explicit %>
<%
'###############################################
' PageName : tenclass_insert.asp
' Discription : 모바일 tenclass
' History : 2018-02-27 이종화
'###############################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/classes/mobile/tenclass_Cls.asp" -->
<%
Dim idx , isusing , mode
Dim srcSDT , srcEDT
Dim mainStartDate, mainEndDate , lp , ii
Dim sDt, sTm, eDt, eTm , gubun , prevDate
Dim mainimg , maincopy , subcopy , adminnotice , mainimage
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
	dim tenClassList
	set tenClassList = new tenClass
	tenClassList.FRectIdx = idx
	tenClassList.GetOneContents()

	maincopy		=	tenClassList.FOneItem.Fmaincopy
	subcopy			=	tenClassList.FOneItem.Fsubcopy
	mainStartDate	=	tenClassList.FOneItem.Fstartdate
	mainEndDate		=	tenClassList.FOneItem.Fenddate
	isusing			=	tenClassList.FOneItem.Fisusing
	adminnotice		=	tenClassList.FOneItem.Fadminnotice
	mainimage		=	tenClassList.FOneItem.Fmainimage

	set tenClassList = Nothing
End If

Dim oSubItemList
set oSubItemList = new tenClass
	oSubItemList.FPageSize = 100
	oSubItemList.FRectidx = idx
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

		if (frm.sTm.value.length != 8) {
			alert("시간을 정확히 입력하세요");
			frm.sTm.focus();
			return;
		}

		if (frm.eTm.value.length != 8) {
			alert("시간을 정확히 입력하세요");
			frm.eTm.focus();
			return;
		}

		if (confirm('저장 하시겠습니까?')){
			//frm.target = "blank";
			frm.submit();
		}
	}
	function jsgolist(){
		self.location.href="/admin/mobile/tenclass/";
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
			$(this).find("input[name^='sort']").each(function(){
				if(i>$(this).val()) i=$(this).val()
			});
			if(i<=0) i=1;
			$(this).find("input[name^='sort']").each(function(){
				$(this).val(i);
				i++;
			});
		}
	});
});

// 상품검색 일괄 등록 (구버전)
function popRegSearchItem() {
<% if idx <> "" then %>
    var popwin = window.open("/admin/itemmaster/pop_itemAddInfo.asp?sellyn=Y&usingyn=Y&defaultmargin=0&acURL=/admin/mobile/tenclass/doSubRegItemCdArray.asp?idx=<%=idx%>", "popup_item", "width=800,height=500,scrollbars=yes,resizable=yes");
    popwin.focus();
<% else %>
	alert("템플릿 컨텐츠 정보를 먼저 등록해주세요.");
<% end if %>
}

// 상품코드 일괄 등록
function popRegArrayItem() {
<% if idx<>"" then %>
    var popwin = window.open('popSubRegItemCdArray.asp?idx=<%=idx%>','popRegArray','width=600,height=300,scrollbars=yes,resizable=yes');
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

//'아이템 삭제
function itemdel(v){
	if (confirm("상품이 삭제됩니다 삭제 하시겠습니까?")){
		document.frmdel.chkIdx.value = v;
		document.frmdel.mode.value = "itemdel";
		document.frmdel.action="doListModify.asp";
		document.frmdel.submit();
	}
}
</script>
<form name="frmdel" method="POST" action="">
<input type="hidden" name="mode" />
<input type="hidden" name="chkIdx" />
</form>
<form name="frm" method="post" action="<%=uploadUrl%>/linkweb/mobile/tenclass_proc.asp" enctype="multipart/form-data" style="margin:0px;">
<table width="900" cellpadding="2" cellspacing="1" class="a" bgcolor="#3d3d3d">
<input type="hidden" name="mode" value="<%=mode%>" />
<input type="hidden" name="adminid" value="<%=session("ssBctId")%>" />
<input type="hidden" name="idx" value="<%=idx%>" />
<input type="hidden" name="prevDate" value="<%=prevDate%>" />
<input type="hidden" name="menupos" value="<%=menupos%>" />
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
	<td bgcolor="#FFF999" align="center" width="10%">메인 이미지</td>
	<td>
		<input type="file" name="mainimage" class="file" title="이벤트 #1" require="N" style="width:50%;" />
		<% if mainimage<>"" then %>
		<br>
		<img src="<%= mainimage %>" width="200" /><br><%= mainimage %>
		<% end if %>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
    <td bgcolor="#FFF999" align="center" width="15%">메인카피</td>
    <td>
		<input type="text" name="maincopy" size="50" value="<%=maincopy%>" />
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td bgcolor="#FFF999" align="center" width="15%">서브카피</td>
    <td>
		<input type="text" name="subcopy" size="80" value="<%=subcopy%>" />
    </td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center">작업자 지시사항</td>
	<td colspan="3"><textarea name="adminnotice" cols="80" rows="8"/><%=adminnotice%></textarea></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center">사용여부</td>
	<td><div style="float:left;"><input type="radio" name="isusing" value="1" <%=chkiif(isusing = "1","checked","")%> checked />사용함 &nbsp;&nbsp;&nbsp; <input type="radio" name="isusing" value="0"  <%=chkiif(isusing = "0","checked","")%>/>사용안함</div> <div style="float:right;margin-top:5px;margin-right:10px;"></div></td>
</tr>
<tr bgcolor="#FFFFFF" align="center">
    <td colspan="2">
		<input type="button" value=" 취 소 " onClick="jsgolist();"/><input type="button" value=" 저 장 " onClick="jsSubmit();"/>
	</td>
</tr>
</table>
</form>

<%
	If idx <> "" then
%>
<p><b>▶ 클래스 정보</b></p>
<!-- // 등록된 소재 목록 --------->
<form name="frmList" method="POST" action="" style="margin:0;">
<input type="hidden" name="mode" value="sub">
<table width="900" align="left" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr bgcolor="#FFFFFF">
		<td colspan="8">
			<table width="100%" border="0" class="a" cellpadding="0" cellspacing="0">
			<tr>
				<td align="left">
					총 <%=oSubItemList.FTotalCount%> 건 /
					<input type="button" value="전체선택" class="button" onClick="chkAllItem()">
					<input type="button" value="상태저장" class="button" onClick="saveList()" title="표시순서 및 사용여부를 일괄저장합니다.">
				</td>
			</tr>
			</table>
		</td>
	</tr>
	<col width="30" />
	<col width="30" />
	<col width="30" />
	<col width="150" />
	<col width="30" />
	<col width="80" />
	<col width="30" />
	<tr align="center" bgcolor="#DDDDFF">
		<td>&nbsp;</td>
		<td>이미지</td>
		<td>상품코드</td>
		<td>상품명</td>
		<td>표시순서</td>
		<td>사용여부</td>
		<td>상품삭제</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td colspan="8">
			<table width="100%" border="0" class="a" cellpadding="0" cellspacing="0">
			<tr>
				<td align="right">
					<input type="button" value="상품코드로 등록" class="button" onClick="popRegArrayItem()" />
					<input type="button" value="상품 추가" class="button" onClick="popRegSearchItem()" />
				</td>
			</tr>
			</table>
		</td>
	</tr>
	<tbody id="subList">
	<% For lp=0 to oSubItemList.FResultCount-1 %>
	<tr align="center" bgcolor="<%=chkIIF(oSubItemList.FItemList(lp).FIsUsing,"#FFFFFF","#F3F3F3")%>#FFFFFF">
		<td><input type="checkbox" name="chkIdx" value="<%=oSubItemList.FItemList(lp).Fdidx%>" /></td>
		<td>
		<%
			if Not(oSubItemList.FItemList(lp).FsmallImage="" or isNull(oSubItemList.FItemList(lp).FsmallImage)) then
				Response.Write "<img src='" & oSubItemList.FItemList(lp).FsmallImage & "' height='50' />"
			end if
		%>
		</td>
		<td>
		<%
			if Not(oSubItemList.FItemList(lp).FItemid="0" or isNull(oSubItemList.FItemList(lp).FItemid) or oSubItemList.FItemList(lp).FItemid="") then
				Response.Write "<input type='text' value='" & oSubItemList.FItemList(lp).FItemid & "' readonly size='5'/>"
			end if
		%>
		</td>
		<td><input type="text" name="itemname<%=oSubItemList.FItemList(lp).Fdidx%>" value="<%=oSubItemList.FItemList(lp).Fitemname%>" size="40"></td>
		<td><input type="text" name="sort<%=oSubItemList.FItemList(lp).Fdidx%>" size="3" class="text" value="<%=oSubItemList.FItemList(lp).Fsortno%>" style="text-align:center;" /></td>
		<td>
			<span class="rdoUsing">
			<input type="radio" name="use<%=oSubItemList.FItemList(lp).Fdidx%>" id="rdoUsing<%=lp%>_1" value="1" <%=chkIIF(oSubItemList.FItemList(lp).FIsUsing,"checked","")%> /><label for="rdoUsing<%=lp%>_1">사용</label>
			<input type="radio" name="use<%=oSubItemList.FItemList(lp).Fdidx%>" id="rdoUsing<%=lp%>_2" value="0" <%=chkIIF(oSubItemList.FItemList(lp).FIsUsing,"","checked")%> /><label for="rdoUsing<%=lp%>_2">삭제</label>
			</span>
		</td>
		<td><input type="button" value="상품삭제" onclick="itemdel('<%=oSubItemList.FItemList(lp).Fdidx%>');"/></td>
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

<%@ language=vbscript %>
<% option explicit %>
<%
'###############################################
' PageName : exhibition_insert.asp
' Discription : 모바일 exhibition
' History : 2016.04.07 이종화
'###############################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/classes/mobile/exhibitionCls.asp" -->
<%
Dim idx , subImage1 , subImage2 , subImage3 , isusing , mode
Dim srcSDT , srcEDT
Dim mainStartDate, mainEndDate , lp , ii
Dim sDt, sTm, eDt, eTm , exhibitiontitle , prevDate , topview , linkurl
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
	dim exhibitionList
	set exhibitionList = new Cexhibition
	exhibitionList.FRectIdx = idx
	exhibitionList.GetOneContents()

	exhibitiontitle	=	exhibitionList.FOneItem.Fexhibitiontitle
	mainStartDate	=	exhibitionList.FOneItem.Fstartdate
	mainEndDate		=	exhibitionList.FOneItem.Fenddate
	isusing			=	exhibitionList.FOneItem.Fisusing
	topview			=	exhibitionList.FOneItem.Ftopview
	linkurl			=	exhibitionList.FOneItem.Flinkurl

	set exhibitionList = Nothing
End If

If Isnull(topview) Or topview = "0" Or  topview = "" Then topview = "N"

Dim oSubItemList
set oSubItemList = new Cexhibition
	oSubItemList.FPageSize = 100
	oSubItemList.FRectlistIdx = idx
	If idx <> "" then
		oSubItemList.GetContentsItemList()
	End If


dim dateOption
dateOption = request("dateoption")

if Not(mainStartDate="" or isNull(mainStartDate)) then
	sDt = left(mainStartDate,10)
	sTm = Num2Str(hour(mainStartDate),2,"0","R")
else
	if srcSDT<>"" then
		sDt = left(srcSDT,10)
	elseif dateOption <> "" then
		sDt = dateOption
	else
		sDt = date
	end if
	sTm = "00"
end if

if Not(mainEndDate="" or isNull(mainEndDate)) then
	eDt = left(mainEndDate,10)
	eTm = Num2Str(hour(mainEndDate),2,"0","R")
else
	if srcEDT<>"" then
		eDt = left(srcEDT,10)
	elseif dateOption <> "" then
		eDt = dateOption
	else
		eDt = date
	end if
	eTm = "23"
end if
%>
<link href="/js/jqueryui/css/jquery-ui.css" rel="stylesheet">
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/jqueryui/jquery-ui-1.10.2.custom.min.js"></script>
<script type='text/javascript'>
	function jsSubmit(){
		var frm = document.frm;

		if (frm.sTm.value.length != 2) {
			alert("시간을 정확히 입력하세요");
			frm.sTm.focus();
			return;
		}

		if (frm.eTm.value.length != 2) {
			alert("시간을 정확히 입력하세요");
			frm.eTm.focus();
			return;
		}

		if (frm.linkurl.value.indexOf("카테고리") > 0 || frm.linkurl.value.indexOf("이벤트번호") > 0 || frm.linkurl.value.indexOf("상품코드") > 0 || frm.linkurl.value.indexOf("브랜드아이디") > 0){
			alert("링크 값을 확인 해주세요.");
			frm.linkurl.focus();
			return;
		}

		if (confirm('저장 하시겠습니까?')){
			//frm.target = "blank";
			frm.submit();
		}
	}
	function jsgolist(){
		self.location.href="/admin/mobile/exhibition/";
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


	$( "#subList").sortable({
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
    var popwin = window.open("/admin/itemmaster/pop_itemAddInfo.asp?sellyn=Y&usingyn=Y&defaultmargin=0&acURL=/admin/mobile/exhibition/doSubRegItemCdArray.asp?listidx=<%=idx%>", "popup_item", "width=800,height=500,scrollbars=yes,resizable=yes");
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

//'아이템 삭제
function itemdel(v){
	if (confirm("상품이 삭제됩니다 삭제 하시겠습니까?")){
		document.frmdel.chkIdx.value = v;
		document.frmdel.mode.value = "itemdel";
		document.frmdel.action="doListModify.asp";
		document.frmdel.submit();
	}
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

</script>
<form name="frmdel" method="POST" action="">
<input type="hidden" name="mode" />
<input type="hidden" name="chkIdx" />
</form>
<table width="900" cellpadding="2" cellspacing="1" class="a" bgcolor="#3d3d3d">
<form name="frm" method="post" action="doexhibition.asp">
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
		<input type="text" name="sTm" size="2" value="<%=sTm%>" maxlength="2"/>:00:00 ~
		<input type="text" id="eDt" name="EndDate" size="10" value="<%=eDt%>" />
		<input type="text" name="eTm" size="2" value="<%=eTm%>" maxlength="2"/>:59:59
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td bgcolor="#FFF999" align="center" width="15%">제목</td>
    <td>
		<input type="text" name="exhibitiontitle" size="50" value="<%If idx="" then%><%=sDt%>&nbsp;시작 메인 기획전 입니다.<% Else %><%=exhibitiontitle%><% End if%>" />
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td bgcolor="#FFF999" align="center" width="15%">더보기 링크</td>
    <td>
		<input type="text" name="linkurl" size="50" value="<%=linkurl%>" /><br/><br/>
		ex)- <span style="cursor:pointer" onClick="putLinkText('event')">이벤트 : /event/eventmain.asp?eventid=<font color="darkred">이벤트코드</font></span><br>
		- <span style="cursor:pointer" onClick="putLinkText('itemid')">상품코드 : /category/category_itemprd.asp?itemid=<font color="darkred">상품코드 (O)</font></span><br>
		- <span style="cursor:pointer" onClick="putLinkText('category')">카테고리 : /category/category_list.asp?disp=<font color="darkred">카테고리</font></span><br>
		- <span style="cursor:pointer" onClick="putLinkText('brand')">브랜드아이디 : /street/street_brand.asp?makerid=<font color="darkred">브랜드아이디</font></span><br/>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center">사용여부</td>
	<td><div style="float:left;"><input type="radio" name="isusing" value="Y" <%=chkiif(isusing = "Y","checked","")%> checked />사용함 &nbsp;&nbsp;&nbsp; <input type="radio" name="isusing" value="N"  <%=chkiif(isusing = "N","checked","")%>/>사용안함</div> <div style="float:right;margin-top:5px;margin-right:10px;"></div></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center">긴급노출</td>
	<td><div style="float:left;"><input type="radio" name="topview" value="Y" <%=chkiif(topview = "Y","checked","")%> />사용함 &nbsp;&nbsp;&nbsp; <input type="radio" name="topview" value="N"  <%=chkiif(topview = "N","checked","")%>/>사용안함</div> <div style="float:right;margin-top:5px;margin-right:10px;"></div></td>
</tr>
<tr bgcolor="#FFFFFF" align="center">
    <td colspan="2">
		<input type="button" value=" 취 소 " onClick="jsgolist();"/><input type="button" value=" 저 장 " onClick="jsSubmit();"/>
	</td>
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
		    	<img src="/images/icon_new_registration.gif" border="0" onclick="popSubEdit('');" style="cursor:pointer;" align="absmiddle">
		    </td>
		</tr>
		</table>
	</td>
</tr>
<col width="30" />
<col width="30" />
<col width="30" />
<col span="3" width="0*" />
<col width="30" />
<col width="30" />
<col width="30" />
<col width="30" />
<tr align="center" bgcolor="#DDDDFF">
    <td>&nbsp;</td>
    <td>소재번호</td>
    <td>이미지</td>
    <td>상품코드</td>
    <td>상품명</td>
    <td>표시순서</td>
    <td>사용여부</td>
    <td>상품삭제</td>
</tr>

<tbody id="subList<%=ii%>">
<% For lp=0 to oSubItemList.FResultCount-1 %>
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
    		Response.Write "<input type='text' value='" & oSubItemList.FItemList(lp).FItemid & "' readonly size='5'/>"
    	end if
    %>
    </td>
	<td><input type="text" name="itemname<%=oSubItemList.FItemList(lp).FsubIdx%>" value="<%=oSubItemList.FItemList(lp).Fitemname%>" size="60"></td>
    <td><input type="text" name="sort<%=oSubItemList.FItemList(lp).FsubIdx%>" size="3" class="text" value="<%=oSubItemList.FItemList(lp).Fsortnum%>" style="text-align:center;" /></td>
    <td>
		<span class="rdoUsing">
		<input type="radio" name="use<%=oSubItemList.FItemList(lp).FsubIdx%>" id="rdoUsing<%=lp%>_1" value="Y" <%=chkIIF(oSubItemList.FItemList(lp).FIsUsing="Y","checked","")%> /><label for="rdoUsing<%=lp%>_1">사용</label><input type="radio" name="use<%=oSubItemList.FItemList(lp).FsubIdx%>" id="rdoUsing<%=lp%>_2" value="N" <%=chkIIF(oSubItemList.FItemList(lp).FIsUsing="N","checked","")%> /><label for="rdoUsing<%=lp%>_2">삭제</label>
		</span>
    </td>
	<td><input type="button" value="상품삭제" onclick="itemdel('<%=oSubItemList.FItemList(lp).FsubIdx%>');"/></td>
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

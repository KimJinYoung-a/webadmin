<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/classes/sitemasterclass/look2018.asp" -->
<%
Dim idx, copyimageurl, bgcolor, mainStartDate, mainEndDate, isusing, adminid, lastadminid, regdate, lastupdate, username, username2, orderby
Dim srcSDT , srcEDT 
Dim lp
Dim sDt, sTm, eDt, eTm , gubun , prevDate
Dim extraurl
Dim paramisusing, mode

	idx = requestCheckvar(request("idx"),16)
	srcSDT = request("sDt")
	srcEDT = request("eDt")
	prevDate = request("prevDate")
	paramisusing = request("paramisusing")


If idx = "" Then 
	mode = "add" 
Else 
	mode = "modify" 
End If 

If idx <> "" then
	dim lookList
	set lookList = new Clook
	lookList.FRectIdx = idx
	lookList.GetOneContents()

	idx			= lookList.FOneItem.Fidx
	copyimageurl	= lookList.FOneItem.Fcopyimageurl
	bgcolor		= lookList.FOneItem.Fbgcolor
	orderby		= lookList.FOneItem.Forderby
	mainStartDate	=	lookList.FOneItem.Fstartdate '// 시작일
	mainEndDate		=	lookList.FOneItem.Fenddate '// 종료일
	isusing			=	lookList.FOneItem.Fisusing '// 사용여부

	set lookList = Nothing
End If 

Dim oSubItemList
set oSubItemList = new Clook
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
		if prevDate = "" then 
			prevDate = sDt
		end if 
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
		if prevDate = "" then 
			prevDate = eDt
		end if 
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

		if (!frm.copyimageurl.value){
			alert("메인카피 이미지를 등록해주세요.");
			return;
		}
		if (!frm.bgcolor.value){
			alert("배경색을 입력해주세요.");
			frm.bgcolor.focus();
			return;
		}
		if (!frm.orderby.value){
			alert("정렬번호를 입력해주세요.");
			frm.orderby.focus();
			return;
		}
		if (!frm.isusing[0].checked && !frm.isusing[1].checked)
		{
			alert("사용여부를 선택하세요!")
			return false;
		}

		if (confirm('저장 하시겠습니까?')){
			//frm.target = "blank";
			frm.submit();
		}
	}
	function jsgolist(){
		self.location.href="/admin/sitemaster/look/?menupos=<%=request("menupos")%>&isusing=<%=paramisusing%>";
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

//소재
function popSubEdit(subidx) {
<% if idx <>"" then %>
    var popwin = window.open('pop_LookItemAddInfo.asp?menupos=<%=menupos%>&prevDate=<%=prevDate%>&paramisusing=<%=paramisusing%>&usingyn=<%=trim(isusing)%>&idx=<%=idx%>&subidx='+subidx,'popTemplateManage','width=800,height=500,scrollbars=yes,resizable=yes');
    popwin.focus();
<% else %>
	alert("템플릿 컨텐츠 정보를 먼저 등록해주세요.");
<% end if %>
}

// 노출상품등록
function popRegSearchItem() {
<% if idx <> "" then %>
    var popwin = window.open("pop_LookItemAddInfo.asp?menupos=<%=menupos%>&prevDate=<%=prevDate%>&paramisusing=<%=paramisusing%>&usingyn=Y&idx=<%=idx%>", "popup_item", "width=800,height=500,scrollbars=yes,resizable=yes");
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

function putLinkText(key) {
	var frm = document.frm;
	switch(key) {
		case 'event':
			frm.extraurl.value='/event/eventmain.asp?eventid=이벤트번호';
			break;
		case 'itemid':
			frm.extraurl.value='/shopping/category_prd.asp?itemid=상품코드';
			break;
	}
}

//-- jsImgView : 이미지 확대화면 새창으로 보여주기 --//
function jsImgView(sImgUrl){
 var wImgView;
 wImgView = window.open('/admin/eventmanage/common/pop_event_detailImg.asp?sUrl='+sImgUrl,'pImg','width=100,height=100');
 wImgView.focus();
}


function jsSetImg(sFolder, sImg, sName, sSpan){ 
	var winImg;
	winImg = window.open('/admin/eventmanage/common/pop_event_uploadimgV2.asp?yr=<%=Year(now())%>&sF='+sFolder+'&sImg='+sImg+'&sName='+sName+'&sSpan='+sSpan,'popImg','width=370,height=150');
	winImg.focus();
}

function jsDelImg(sName, sSpan){
	if(confirm("이미지를 삭제하시겠습니까?\n\n삭제 후 저장버튼을 눌러야 처리완료됩니다.")){
	   eval("document.all."+sName).value = "";
	   eval("document.all."+sSpan).style.display = "none";
	}
}
</script>
<form name="frm" method="post" action="dolook.asp" style="margin:0px;">
<input type="hidden" name="mode" value="<%=mode%>">
<input type="hidden" name="adminid" value="<%=session("ssBctId")%>">
<input type="hidden" name="idx" value="<%=idx%>">
<input type="hidden" name="prevDate" value="<%=prevDate%>">
<input type="hidden" name="paramisusing" value="<%=paramisusing%>">
<input type="hidden" name="copyimageurl" value="<%=copyimageurl%>">
<table width="900" cellpadding="2" cellspacing="1" class="a" bgcolor="#3d3d3d">
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
		<input type="text" id="sDt" name="StartDate" size="10" value="<%=chkiif(mode="add",prevDate,sDt)%>" />
		<input type="text" name="sTm" size="8" value="<%=sTm%>" /> ~
		<input type="text" id="eDt" name="EndDate" size="10" value="<%=chkiif(mode="add",prevDate,eDt)%>" />
		<input type="text" name="eTm" size="8" value="<%=eTm%>" />
    </td>
</tr>
<tr bgcolor="#FFFFFF" id="tmpstyle1">
	<td bgcolor="#DDDDFF" align="center" width="15%">메인카피이미지</td>
	<td><input type="button" name="limg" value="메인카피 이미지 등록" onClick="jsSetImg('pcmainlook','<%=copyimageurl%>','copyimageurl','lookmainimg')" class="button">
		<div id="lookmainimg" style="padding: 5 5 5 5">
			<%IF copyimageurl <> "" THEN %>
			<a href="javascript:jsImgView('<%=copyimageurl%>')"><img  src="<%=copyimageurl%>" width="400" border="0"></a>
			<a href="javascript:jsDelImg('copyimageurl','lookmainimg');"><img src="/images/icon_delete2.gif" border="0"></a>
			<%END IF%>
		</div>
		<%=copyimageurl%>
	</td>
</tr>
<tr bgcolor="#FFFFFF" id="tmpstyle2">
	<td bgcolor="#DDDDFF"  align="center" width="15%">배경색</td>
	<td>
		<input type="text" name="bgcolor" value="<%=bgcolor%>" style="width:20%;" /> <font color="red">#빼고 색상코드 넣어주세요.</font>
	</td>
</tr>
<tr bgcolor="#FFFFFF" id="tmpstyle2">
	<td bgcolor="#DDDDFF"  align="center" width="15%">정렬번호</td>
	<td>
		<input type="text" name="orderby" value="<%=orderby%>" style="width:10%;" /> <font color="red">정렬번호는 오름차순 입니다.</font>
	</td>
</tr>

<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center">사용여부</td>
	<td><div style="float:left;"><input type="radio" name="isusing" value="Y" <%=chkiif(Trim(isusing) = "Y","checked","")%> />사용함 &nbsp;&nbsp;&nbsp; <input type="radio" name="isusing" value="N"  <%=chkiif(Trim(isusing) = "N","checked","")%>/>사용안함</div> <div style="float:right;margin-top:5px;margin-right:10px;"></div></td>
</tr>

<tr bgcolor="#FFFFFF" align="center">
    <td colspan="2"><input type="button" value=" 취 소 " onClick="jsgolist();"/><input type="button" value=" 저 장 " onClick="jsSubmit();"/></td>
</tr>
</table>
</form>

<%
	If idx <> "" then
%>

<!-- // 등록된 소재 목록 --------->
<form name="frmList" method="POST" action="" style="margin:0;">
<input type="hidden" name="mode" value="sub">
<table width="900" align="left" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="#FFFFFF">
	<td colspan="10">
		<table width="100%" border="0" class="a" cellpadding="0" cellspacing="0">
		<tr>
		    <td align="left">
		    	총 <%=oSubItemList.FTotalCount%> 건<!-- /
		    	<input type="button" value="전체선택" class="button" onClick="chkAllItem()">
		    	<input type="button" value="상태저장" class="button" onClick="saveList()" title="사용여부를 일괄저장합니다.">
				-->
		    </td>
		    <td align="right">
		    	<!--<input type="button" value="상품코드로 등록" class="button" onClick="popRegArrayItem()" />//-->
		    	<input type="button" value="상품 추가" class="button" onClick="popRegSearchItem()" />
		    	<!--<img src="/images/icon_new_registration.gif" border="0" onclick="popSubEdit('')" style="cursor:pointer;" align="absmiddle">//-->
		    </td>
		</tr>
		</table>
	</td>
</tr>
<col width="70" />
<col width="70" />
<col width="100" />
<col width="100" />
<col width="300" />
<col width="50" />
<col width="100" />
<col width="30" />
<tr align="center" bgcolor="#DDDDFF">
    <td>소재번호</td>
    <td>타입</td>
    <td>이미지</td>
    <td>상품코드</td>
    <td>상품명</td>
    <td>정렬번호</td>
    <td>사용여부</td>
</tr>
<tbody id="subList">
<%	For lp=0 to oSubItemList.FResultCount-1 %>
<tr align="center" bgcolor="<%=chkIIF(Trim(oSubItemList.FItemList(lp).FIsUsing)="Y","#FFFFFF","#F3F3F3")%>#FFFFFF">
    <td onclick="popSubEdit(<%=oSubItemList.FItemList(lp).FsubIdx%>)" style="cursor:pointer;"><%=oSubItemList.FItemList(lp).FsubIdx%></td>
    <td onclick="popSubEdit(<%=oSubItemList.FItemList(lp).FsubIdx%>)" style="cursor:pointer;">
		<%=chkIIF(Trim(oSubItemList.FItemList(lp).Fdisplaytype)="U","위","아래")%>
	</td>
    <td onclick="popSubEdit(<%=oSubItemList.FItemList(lp).FsubIdx%>)" style="cursor:pointer;">
    <%
    	if Not(oSubItemList.FItemList(lp).Fitemimage="" or isNull(oSubItemList.FItemList(lp).Fitemimage)) then
    		Response.Write "<img src='" & oSubItemList.FItemList(lp).Fitemimage & "' height='50' />"
    	end if
    %>
    </td>
    <td>
    <%
    	if Not(oSubItemList.FItemList(lp).FItemid="0" or isNull(oSubItemList.FItemList(lp).FItemid) or oSubItemList.FItemList(lp).FItemid="") then
    		Response.Write oSubItemList.FItemList(lp).FItemid
    	end if
    %>
    </td>
	<td align="left" style="padding-left:5px;">
		<%=oSubItemList.FItemList(lp).Fitemname%><br/>
	</td>
    <td><%=oSubItemList.FItemList(lp).Forderby%></td>
    <td>
		<span class="rdoUsing"><%=chkIIF(Trim(oSubItemList.FItemList(lp).FIsUsing)="Y","사용","사용안함")%>
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
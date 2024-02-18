<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description : PC메인관리 저스트원데이
' History : 서동석 생성
'			2022.07.04 한용민 수정(isms취약점조치)
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/classes/sitemasterclass/just1DayCls2018New.asp" -->
<%
Dim idx , subImage1 , subImage2 , subImage3 , isusing , mode, paramisusing, bannerimage
Dim srcSDT , srcEDT 
Dim mainStartDate, mainEndDate , lp
Dim sDt, sTm, eDt, eTm , gubun , title , prevDate , is1day
Dim linkurl, workertext, vplatform
Dim subtitle, saleper
Dim vType

	idx = requestCheckvar(getNumeric(request("idx")),16)
	srcSDT = request("sDt")
	srcEDT = request("eDt")
	prevDate = request("prevDate")
	paramisusing = request("paramisusing")
	vType = request("type")
	vplatform = "pc"

	If idx = "" Then 
		mode = "add" 
	Else 
		mode = "modify" 
	End If 

	If idx <> "" then
		dim just1dayList
		set just1dayList = new Cjust1Day
		just1dayList.FRectIdx = idx
		just1dayList.GetOneContents()

		title	=	ReplaceBracket(just1dayList.FOneItem.Ftitle) '// 제목(front에 표시되는 경우 없음)
		mainStartDate	=	just1dayList.FOneItem.Fstartdate '// 시작일
		mainEndDate		=	just1dayList.FOneItem.Fenddate '// 종료일
		isusing			=	just1dayList.FOneItem.Fisusing '// 사용여부
		saleper			=	just1dayList.FOneItem.Fsaleper '// 할인율
		vType			=	just1dayList.FOneItem.FType '// type(just1day, event)
		bannerimage		=	just1dayList.FOneItem.FbannerImage '// 기획전용 배너이미지
		linkurl			=	just1dayList.FOneItem.FlinkUrl '// 기획전용 배너 링크url
		workertext		=	just1dayList.FOneItem.FworkerText '// 작업자 전달사항(기획전일 경우)	
		vplatform		=	just1dayList.FOneItem.Fplatform '// 플랫폼(pc,mobile)	

		'// 2019-12-11 디자인팀 요청으로 기존 주말특가(event) Just1Day 부분 이미지 제목 사용하던걸 텍스트로 대체
		'// 해서.. 주말특가(vType=event)일 경우 Front에 표시되지 않는 타이틀을 이벤트 메인카피로,
		'// workertext를 이벤트 서브카피로 사용하고 linkUrl은 webadmin 상에서 이벤트 선택시 자동으로 생성해줌.
		'// 다만 주말특가가 아닌 기존 Just1Day는 그대로감(title Front에 표시 안됨 및 기타 기능 그대로..)

		set just1dayList = Nothing
	End If 

	If Trim(vType)="" Then
		vType = "just1day"
	End If


	Dim oSubItemList
	set oSubItemList = new Cjust1Day
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

		<% if vType="just1day" then %>
			if (!frm.title.value){
				alert("제목을 입력해주세요.");
				frm.title.focus();
				return;
			}

			if (!frm.saleper.value){
				alert("최대할인율을 확인 해주세요.");
				frm.saleper.focus();
				return;
			}
		<% end if %>

		<% if vType="event" then %>		
			if (!frm.linkurl.value){
				alert("기획전을 선택해주세요.");
				frm.linkurl.focus();
				return;
			}

			if (!frm.title.value){
				alert("기획전 메인카피를 입력해주세요.");
				frm.title.focus();
				return;
			}

			if (!frm.workertext.value){
				alert("기획전 서브카피를 입력해주세요.");
				frm.workertext.focus();
				return;
			}

			<%'// 최대 할인율은 입력 안하는 경우도 있다고 해서 체크 안함 %>
			/*
			if (!frm.saleper.value){
				alert("최대할인율을 확인 해주세요.");
				frm.saleper.focus();
				return;
			}
			*/			
		
			if (frm.linkurl.value.indexOf("이벤트번호") > 0 || frm.linkurl.value.indexOf("상품코드") > 0){
				alert("배너 링크 값을 확인 해주세요.");
				frm.linkurl.focus();
				return;
			}
		<% end if %>

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
		self.location.href="/admin/sitemaster/just1day2018/?menupos=<%=request("menupos")%>&isusing=<%=paramisusing%>";
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
    var popwin = window.open('popSubItemEdit.asp?listIdx=<%=idx%>&subIdx='+subidx,'popTemplateManage','width=800,height=500,scrollbars=yes,resizable=yes');
    popwin.focus();
<% else %>
	alert("템플릿 컨텐츠 정보를 먼저 등록해주세요.");
<% end if %>
}

function popRegSearchItem() {
<% if idx <> "" then %>
    var popwin = window.open("/admin/sitemaster/just1day2018/popSubItemEdit.asp?listidx=<%=idx%>", "popup_item", "width=800,height=500,scrollbars=yes,resizable=yes");
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
			frm.linkurl.value='/event/eventmain.asp?eventid=이벤트번호';
			break;
		case 'itemid':
			frm.linkurl.value='/shopping/category_prd.asp?itemid=상품코드';
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

function jsChangeTypeJust1Day(typevalue)
{
	location.href='/admin/sitemaster/just1day2018/just1day_insert.asp?menupos=<%=request("menupos")%>&idx=<%=idx%>&sDt='+document.frm.sDt.value+'&eDt='+document.frm.eDt.value+'&prevDate=<%=prevDate%>&paramisusing=<%=paramisusing%>&type='+typevalue;
}

//-- jsLastEvent : 지난 이벤트 불러오기 --//
function jsLastEvent(){
	winLast = window.open('pop_event_lastlist.asp','pLast','width=800,height=600, scrollbars=yes')
	winLast.focus();
}

function jsCheckJust1DayEventUrl(){
	if(document.frm.linkurl.value=="") {
		alert("기획전을 먼저 선택해주세요.");
		return;
	}
	else {
		<% If LCase(application("Svr_Info"))="dev" Then %>
			window.open('http://2015www.10x10.co.kr'+document.frm.linkurl.value);
		<% ElseIf LCase(application("Svr_Info"))="staging" Then %>
			window.open('http://stgwww.10x10.co.kr'+document.frm.linkurl.value);
		<% Else %>
			window.open('http://www.10x10.co.kr'+document.frm.linkurl.value);
		<% End If %>
	}
}
</script>
<form name="frm" method="post" action="dojust1day.asp" style="margin:0px;">
<input type="hidden" name="mode" value="<%=mode%>">
<input type="hidden" name="adminid" value="<%=session("ssBctId")%>">
<input type="hidden" name="idx" value="<%=idx%>">
<input type="hidden" name="prevDate" value="<%=prevDate%>">
<input type="hidden" name="paramisusing" value="<%=paramisusing%>">
<input type="hidden" name="bannerimage" value="<%=bannerimage%>">
<input type="hidden" name="platform" value="<%=vplatform%>">
<table width="1100" cellpadding="2" cellspacing="1" class="a" bgcolor="#3d3d3d">
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
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center">구분</td>
	<td>
		<div style="float:left;">
			<% If idx="" Then %>
				<input type="radio" name="type" value="just1day" <%=chkiif(vType = "just1day","checked","")%> onclick="jsChangeTypeJust1Day('just1day');"/>JUST 1 DAY
				
				&nbsp;&nbsp;&nbsp; <input type="radio" name="type" value="event"  <%=chkiif(vType = "event","checked","")%> onclick="jsChangeTypeJust1Day('event');"/>기획전
			<% Else %>
				<% If vType="just1day" Then %>
					JUST 1 DAY
					<input type="hidden" name="type" value="just1day">
				<% End If %>
				<% If vType="event" Then %>
					기획전
					<input type="hidden" name="type" value="event">
				<% End If %>
			<% End If %>
		</div>
	</td>
</tr>
<% If vType="just1day" Then %>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center" width="15%">제목</td>
	<td>
		<input type="text" name="title" size="50" value="<%=title%>" /> <font color="red">Front에 표시 안됩니다.</font>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center" width="15%">할인율</td>
	<td>
		<input type="text" name="saleper" size="50" value="<%=saleper%>" /> <font color="red">ex) ~91%</font>
	</td>
</tr>
<% End If %>
<% If vtype="event" Then %>
	<tr bgcolor="#FFFFFF">
		<td bgcolor="#FFF999"  align="center" width="15%">기획전 선택</td>
		<td>
			<input type="text" name="linkurl" value="<%=linkurl%>" size="50" readonly />&nbsp;&nbsp;<a href="" onclick="jsLastEvent();return false;">[기획전 불러오기]</a>&nbsp;&nbsp;<a href="" onclick="jsCheckJust1DayEventUrl();return false;">[기획전 확인하기]</a>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td bgcolor="#FFF999" align="center" width="15%">기획전 메인카피</td>
		<td>
			<input type="text" name="title" size="50" maxlength="30" value="<%=title%>" /> <font color="red">30자 까지 입력 가능합니다.</font>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td bgcolor="#FFF999" align="center" width="15%">기획전 서브카피</td>
		<td><input type="text" name="workertext" size="80" maxlength="55" value="<%=workertext%>"> <font color="red">55자 까지 입력 가능합니다.</font>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td bgcolor="#FFF999" align="center" width="15%">할인율</td>
		<td>
			<input type="text" name="saleper" size="50" value="<%=saleper%>" /> <font color="red">ex) ~91%</font>
		</td>
	</tr>	
<% End If %>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center">사용여부</td>
	<td><div style="float:left;"><input type="radio" name="isusing" value="Y" <%=chkiif(isusing = "Y","checked","")%> />사용함 &nbsp;&nbsp;&nbsp; <input type="radio" name="isusing" value="N"  <%=chkiif(isusing = "N","checked","")%>/>사용안함</div> <div style="float:right;margin-top:5px;margin-right:10px;"></div></td>
</tr>

<tr bgcolor="#FFFFFF" align="center">
    <td colspan="2"><input type="button" value=" 취 소 " onClick="jsgolist();"/><input type="button" value=" 저 장 " onClick="jsSubmit();"/></td>
</tr>
</table>
</form>

<%
	If idx <> "" then
%>
<form name="frmList" method="POST" action="" style="margin:0;">
<input type="hidden" name="mode" value="sub">
<table width="1100" align="left" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="#FFFFFF">
	<td colspan="15">
		<table width="100%" border="0" class="a" cellpadding="0" cellspacing="0">
		<tr>
		    <td align="left">
		    	총 <%=oSubItemList.FTotalCount%> 건 
		    	<!--input type="button" value="전체선택" class="button" onClick="chkAllItem()">
		    	<input type="button" value="상태저장" class="button" onClick="saveList()" title="사용여부를 일괄저장합니다."-->
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
<col width="50" />
<col width="50" />
<col width="50" />
<col width="200" />
<col width="30" />
<col width="80" />
<col width="80" />
<col width="80" />
<col width="50" />
<col width="50" />
<col width="50" />
<col width="50" />
<tr align="center" bgcolor="#DDDDFF">
    <td>구분</td>
    <td>IDX</td>
    <td>상품코드</td>
    <td>노출명</td>
    <td>FrontIMAGE</td>
    <td>가격</td>
    <td>할인율</td>
    <td>정렬순서</td>
    <td>사용여부</td>
    <td></td>
</tr>
<tbody id="subList">
<%	For lp=0 to oSubItemList.FResultCount-1 %>
<tr align="center" bgcolor="<%=chkIIF(oSubItemList.FItemList(lp).FIsUsing="Y","#FFFFFF","#F3F3F3")%>#FFFFFF">
    <td onclick="popSubEdit(<%=oSubItemList.FItemList(lp).FsubIdx%>)" style="cursor:pointer;">
	<%
		If oSubItemList.FItemList(lp).Fitemdiv="21" Then
			response.write "딜상품"
		Else
			response.write "일반상품"
		End If
	%>
	</td>
    <td onclick="popSubEdit(<%=oSubItemList.FItemList(lp).FsubIdx%>)" style="cursor:pointer;"><%=oSubItemList.FItemList(lp).FsubIdx%></td>
    <td onclick="popSubEdit(<%=oSubItemList.FItemList(lp).FsubIdx%>)" style="cursor:pointer;"><%=oSubItemList.FItemList(lp).FItemid%></td>
    <td onclick="popSubEdit(<%=oSubItemList.FItemList(lp).FsubIdx%>)" style="cursor:pointer;"><%=oSubItemList.FItemList(lp).FTitle%></td>
    <td onclick="popSubEdit(<%=oSubItemList.FItemList(lp).FsubIdx%>)" style="cursor:pointer;">
    <%
    	if Not(oSubItemList.FItemList(lp).FitemFrontimage="") then
    		Response.Write "<img src='" & oSubItemList.FItemList(lp).FitemFrontimage & "' height='50' />"
		Else
    		Response.Write "<img src='" & oSubItemList.FItemList(lp).FsmallImage & "' height='50' />"
    	end if
    %>
    </td>
    <td onclick="popSubEdit(<%=oSubItemList.FItemList(lp).FsubIdx%>)" style="cursor:pointer;"><%=oSubItemList.FItemList(lp).FitemPrice%></td>
    <td onclick="popSubEdit(<%=oSubItemList.FItemList(lp).FsubIdx%>)" style="cursor:pointer;"><%=oSubItemList.FItemList(lp).Fitemsaleper%></td>
    <td onclick="popSubEdit(<%=oSubItemList.FItemList(lp).FsubIdx%>)" style="cursor:pointer;"><%=oSubItemList.FItemList(lp).Fsortnum%></td>
    <td onclick="popSubEdit(<%=oSubItemList.FItemList(lp).FsubIdx%>)" style="cursor:pointer;">
		<%
			If Trim(oSubItemList.FItemList(lp).Fisusing="Y") Then 
				Response.write "사용"
			Else
				Response.write "사용안함"
			End If
		%>
	</td>
    <td onclick="popSubEdit(<%=oSubItemList.FItemList(lp).FsubIdx%>)" style="cursor:pointer;"><input type="button" value="수정"></td>
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
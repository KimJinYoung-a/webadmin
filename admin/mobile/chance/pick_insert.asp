<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/classes/mobile/pickCls.asp" -->
<%
Dim idx , subImage1 , subImage2 , subImage3 , isusing , mode, paramisusing
Dim srcSDT , srcEDT 
Dim mainStartDate, mainEndDate , lp
Dim sDt, sTm, eDt, eTm , gubun , picktitle , prevDate , is1day
Dim extraurl
Dim subtitle , saleper

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
	dim pickList
	set pickList = new Cpick
	pickList.FRectIdx = idx
	pickList.GetOneContents()

	gubun			=	pickList.FOneItem.Fgubun
	picktitle		=	pickList.FOneItem.Ftitle
	mainStartDate	=	pickList.FOneItem.Fstartdate
	mainEndDate		=	pickList.FOneItem.Fenddate 
	isusing			=	pickList.FOneItem.Fisusing
	is1day			=	pickList.FOneItem.Fis1day

	subImage1		=	pickList.FOneItem.FsubImage1
	extraurl		=	pickList.FOneItem.Fextraurl
	subtitle		=	pickList.FOneItem.Fsubtitle
	saleper			=	pickList.FOneItem.Fsaleper

	set pickList = Nothing
End If 

Dim oSubItemList
set oSubItemList = new Cpick
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

		if (!frm.isusing[0].checked && !frm.isusing[1].checked)
		{
			alert("사용여부를 선택하세요!")
			return false;
		}

		if (frm.extraurl.value.indexOf("이벤트번호") > 0 || frm.extraurl.value.indexOf("상품코드") > 0){
			alert("링크 값을 확인 해주세요.");
			frm.extraurl.focus();
			return;
		}
		
		// 주말 등록 일때만 사용
		if (frm.is1day[1].checked){
			if (!frm.saleper.value){
				alert("할인율을 확인 해주세요.");
				frm.saleper.focus();
				return;
			}
		}

		if (confirm('저장 하시겠습니까?')){
			//frm.target = "blank";
			frm.submit();
		}
	}
	function jsgolist(){
		self.location.href="/admin/mobile/chance/?menupos=<%=request("menupos")%>&isusing=<%=paramisusing%>";
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
    var popwin = window.open('popSubItemEdit.asp?listIdx=<%=idx%>&subIdx='+subidx,'popTemplateManage','width=800,height=300,scrollbars=yes,resizable=yes');
    popwin.focus();
<% else %>
	alert("템플릿 컨텐츠 정보를 먼저 등록해주세요.");
<% end if %>
}

// 상품검색 일괄 등록
function popRegSearchItem() {
<% if idx <> "" then %>
	var is1day=document.frm.is1day.value;
	if(is1day=="Y")
	{
		is1day="just1day"
	}
    var popwin = window.open("/admin/itemmaster/pop_itemAddInfo.asp?sellyn=Y&usingyn=Y&defaultmargin=0&acURL=/admin/mobile/chance/doSubRegItemCdArray.asp?listidx=<%=idx%>&ptype="+is1day, "popup_item", "width=800,height=500,scrollbars=yes,resizable=yes");
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
			frm.extraurl.value='/category/category_itemprd.asp?itemid=상품코드';
			break;
	}
}

function chktmp(v){
	if (v == 'c'){
		document.getElementById("tmpstyle1").style.display = "";
		document.getElementById("imagetitle").textContent = "휴일/주말이미지";
		document.getElementsByName("wimg")[0].value = "휴일/주말이미지 등록";
		document.getElementById("tmpstyle2").style.display = "";
		document.getElementById("weektitle").style.display = "none";
		document.getElementById("subtitle").style.display = "none";
	}else if (v == 'b'){
		document.getElementById("weektitle").style.display = "";
		document.getElementById("subtitle").style.display = "";
		document.getElementById("tmpstyle1").style.display = "none";
		document.getElementById("tmpstyle2").style.display = "";
	}else{
		document.getElementById("tmpstyle1").style.display = "";
		document.getElementById("imagetitle").textContent = "상품이미지";
		document.getElementsByName("wimg")[0].value = "상품이미지 등록";
		document.getElementById("tmpstyle2").style.display = "none";
		document.getElementById("weektitle").style.display = "none";
		document.getElementById("subtitle").style.display = "none";
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
<form name="frm" method="post" action="dopick.asp" style="margin:0px;">
<input type="hidden" name="mode" value="<%=mode%>">
<input type="hidden" name="adminid" value="<%=session("ssBctId")%>">
<input type="hidden" name="idx" value="<%=idx%>">
<input type="hidden" name="prevDate" value="<%=prevDate%>">
<input type="hidden" name="paramisusing" value="<%=paramisusing%>">
<input type="hidden" name="todayban" value="<%=subImage1%>">
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
		<input type="text" id="sDt" name="StartDate" size="10" value="<%=sDt%>" />
		<input type="text" name="sTm" size="8" value="<%=sTm%>" /> ~
		<input type="text" id="eDt" name="EndDate" size="10" value="<%=eDt%>" />
		<input type="text" name="eTm" size="8" value="<%=eTm%>" />
    </td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center">JUST 1 DAY 여부</td>
	<td><div style="float:left;"><input type="radio" name="is1day" value="Y" <%=chkiif(is1day = "Y","checked","")%> checked onclick="chktmp('a');"/>JUST 1 DAY &nbsp;&nbsp;&nbsp; <input type="radio" name="is1day" value="W"  <%=chkiif(is1day = "W","checked","")%> onclick="chktmp('b');"/>주말특가(기획전) &nbsp;&nbsp;&nbsp;<input type="radio" name="is1day" value="N"  <%=chkiif(is1day = "N","checked","")%> onclick="chktmp('c');"/>이미지배너(기존 주말배너)</div> <div style="float:right;margin-top:5px;margin-right:10px;"></div></td>
</tr>
<tr bgcolor="#FFFFFF">
    <td bgcolor="#FFF999" align="center" width="15%">제목</td>
    <td>
		<input type="text" name="title" size="50" value="<%=picktitle%>" /> <span id="weektitle" style="display:<%=chkiif(is1day="W","","none")%>;"><font color="red">한글로만 입력 예시) 주/말/특/가</font></span>
		<span id="subtitle" style="display:<%=chkiif(is1day="W","","none")%>;">&nbsp; 할인율 : <input type="text" name="saleper" size="20" value="<%=saleper%>" placeholder="예시) 40%~"/></span>
    </td>
</tr>
<tr bgcolor="#FFFFFF" id="tmpstyle1" style="display:<%=chkiif(is1day="Y" Or is1day="N","","none")%>;">
	<td bgcolor="#DDDDFF" align="center" width="15%" id="imagetitle"><%=chkiif(is1day="Y","상품이미지","휴일/주말이미지")%></td>
	<td><input type="button" name="wimg" value="<%=chkiif(is1day="Y","상품이미지","휴일/주말이미지")%> 등록" onClick="jsSetImg('today','<%=subImage1%>','todayban','weekendimg')" class="button">
		<div id="weekendimg" style="padding: 5 5 5 5">
			<%IF subImage1 <> "" THEN %>
			<a href="javascript:jsImgView('<%=subImage1%>')"><img  src="<%=subImage1%>" width="400" border="0"></a>
			<a href="javascript:jsDelImg('todayban','weekendimg');"><img src="/images/icon_delete2.gif" border="0"></a>
			<%END IF%>
		</div>
		<%=subImage1%>
	</td>
</tr>
<tr bgcolor="#FFFFFF" id="tmpstyle2" style="display:<%=chkiif(is1day="N" Or is1day="W","","none")%>;">
	<td bgcolor="#DDDDFF"  align="center" width="15%">주말특가 / 이미지배너 링크</td>
	<td>
		<input type="text" name="extraurl" value="<%=extraurl%>" style="width:80%;" /><br/>
		- <font color="red"><strong>app & mobile 공용</strong></font> - <br/>
		- <span style="cursor:pointer" onClick="putLinkText('event')">이벤트 링크 : /event/eventmain.asp?eventid=<font color="darkred">이벤트코드</font></span><br>
		- <span style="cursor:pointer" onClick="putLinkText('itemid')">상품코드 링크 : /category/category_itemprd.asp?itemid=<font color="darkred">상품코드 (O)</font></span><br>
	</td>
</tr>
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
<p><b>▶ 소재 정보</b></p>
<p>
	<strong>
		※ [ON]상품관리 >> 상품쿠폰관리 대상 상품 팝업 클릭 >> 상품번호 하단 URL생성 >> 모바일웹 링크 참고<br/>예) http://m.10x10.co.kr/category/category_itemprd.asp?itemid=663507&<span style="color:blue">ldv=<span style="color:red">MzIwMCAg</span></span>
	</strong>
</p>
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
		    	<input type="button" value="상태저장" class="button" onClick="saveList()" title="사용여부를 일괄저장합니다.">
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
    <td>상품코드 // ldv</td>
    <td>상품명</td>
    <td>정렬번호</td>
    <td>라벨</td>
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
    		Response.Write "[<input type='text' value='" & oSubItemList.FItemList(lp).FItemid & "' size=5 />]"
    	end if
    %>
	&nbsp;ldv=<input type="text" name="ldv<%=oSubItemList.FItemList(lp).FsubIdx%>" value="<%=oSubItemList.FItemList(lp).Fldv%>"/ size="5">
    </td>
	<td align="left" style="padding-left:5px;">
		상품명 : <input type="text" name="itemname<%=oSubItemList.FItemList(lp).FsubIdx%>" value="<%=oSubItemList.FItemList(lp).Fitemname%>"/ size="35"><br/>
	</td>
    <td><input type="text" name="sort<%=oSubItemList.FItemList(lp).FsubIdx%>" size="3" class="text" value="<%=oSubItemList.FItemList(lp).Fsortnum%>" style="text-align:center;" /></td>
    <td>
		<select name="label<%=oSubItemList.FItemList(lp).FsubIdx%>">
			<option value="0" <%=chkiif(oSubItemList.FItemList(lp).Flabel="","selected","")%>>선택</option>
			<option value="1" <%=chkiif(oSubItemList.FItemList(lp).Flabel="1","selected","")%>>10x10 ONLY</option>
			<option value="2" <%=chkiif(oSubItemList.FItemList(lp).Flabel="2","selected","")%>>HOT ITEM</option>
			<option value="3" <%=chkiif(oSubItemList.FItemList(lp).Flabel="3","selected","")%>>WISH NO.1</option>
			<option value="4" <%=chkiif(oSubItemList.FItemList(lp).Flabel="4","selected","")%>>BEST ITEM</option>
			<option value="5" <%=chkiif(oSubItemList.FItemList(lp).Flabel="5","selected","")%>>1Day</option>
			<option value="6" <%=chkiif(oSubItemList.FItemList(lp).Flabel="6","selected","")%>>한정</option>
			<option value="7" <%=chkiif(oSubItemList.FItemList(lp).Flabel="7","selected","")%>>1+1</option>
			<option value="8" <%=chkiif(oSubItemList.FItemList(lp).Flabel="8","selected","")%>>할인</option>
			<option value="9" <%=chkiif(oSubItemList.FItemList(lp).Flabel="9","selected","")%>>무료배송</option>
		</select>
	</td>
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
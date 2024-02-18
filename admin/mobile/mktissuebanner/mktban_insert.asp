<%@ language=vbscript %>
<% option explicit %>
<%
'###############################################
' PageName : mktban_insert.asp
' Discription : 모바일 mktbanner_new
' History : 2015-01-07 이종화
'###############################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/classes/mobile/mktevtbannerCls.asp" -->
<%
'###############################################
'이벤트 신규 등록시
'###############################################
Dim eCode
Dim idx , isusing , mode , topfixed, evtgubun
Dim srcSDT , srcEDT 
Dim mainStartDate, mainEndDate
Dim sDt, sTm, eDt, eTm
Dim a_eventid , m_eventid
Dim prevDate 
Dim stdt , eddt , sortnum , mktimg
Dim gubun , altname

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
	dim oMktbannerOne
	set oMktbannerOne = new CEvtMktbanner
	oMktbannerOne.FRectIdx = idx
	oMktbannerOne.GetOneContents()

	idx					=	oMktbannerOne.FOneItem.Fidx			
	mktimg				=	oMktbannerOne.FOneItem.Fmktimg		
	a_eventid			=	oMktbannerOne.FOneItem.Fa_eventid
	m_eventid			=	oMktbannerOne.FOneItem.Fm_eventid
	mainStartDate		=	oMktbannerOne.FOneItem.Fstartdate		
	mainEndDate			=	oMktbannerOne.FOneItem.Fenddate		
	isusing				=	oMktbannerOne.FOneItem.Fisusing		
	sortnum				=	oMktbannerOne.FOneItem.Fsortnum		
	gubun				=	oMktbannerOne.FOneItem.Fgubun		
	altname				=	oMktbannerOne.FOneItem.Faltname
	topfixed			=	oMktbannerOne.FOneItem.Ftopfixed
	evtgubun			=	oMktbannerOne.FOneItem.Fevtgubun	''2016-04-28 유태욱 추가(기획전1, 마케팅2 이벤트 구분)

	set oMktbannerOne = Nothing
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
	function jsSubmit(){
		var frm = document.frm;
		var gubun = document.getElementById("gubun");

		if ( gubun.selectedIndex == 0 ){
			alert("구분을 선택 해주세요");
			gubun.focus();
			return;
		}else{
			if ( gubun.value == "1" ){
				if (!frm.m_eventid.value){
					alert("모바일 이벤트 코드를 입력 해주세요");
					frm.m_eventid.focus();
					return;
				}
				if (!frm.a_eventid.value){
					alert("앱 이벤트 코드를 입력 해주세요");
					frm.a_eventid.focus();
					return;
				}
			}else if ( gubun.value == "2" ){
				if (!frm.m_eventid.value){
					alert("모바일 이벤트 코드를 입력 해주세요");
					frm.m_eventid.focus();
					return;
				}
			}else if ( gubun.value == "3" ){
				if (!frm.a_eventid.value){
					alert("앱 이벤트 코드를 입력 해주세요");
					frm.a_eventid.focus();
					return;
				}
			}
		}
		if (confirm('저장 하시겠습니까?')){
			//frm.target = "blank";
			frm.submit();
		}
	}
	function jsgolist(){
		self.close();
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
		if (v == "1"){
			$("#url1").css("display","block");
			$("#url2").css("display","block");
		}else if (v == "2"){
			$("#url1").css("display","block");
			$("#url2").css("display","none");
		}else if (v == "3"){
			$("#url1").css("display","none");
			$("#url2").css("display","block");
		}else{
			$("#url1").css("display","none");
			$("#url2").css("display","none");
		}
	}
</script>
<table width="750" cellpadding="2" cellspacing="1" class="a" bgcolor="#3d3d3d">
<form name="frm" method="post" action="<%=uploadUrl%>/linkweb/mobile/mktbanner_proc.asp" enctype="multipart/form-data" style="margin:0px;">
<input type="hidden" name="mode" value="<%=mode%>">
<input type="hidden" name="adminid" value="<%=session("ssBctId")%>">
<input type="hidden" name="idx" value="<%=idx%>">
<input type="hidden" name="prevDate" value="<%=prevDate%>">
<input type="hidden" name="menupos" value="<%=menupos%>">
<tr bgcolor="#FFFFFF">
    <td bgcolor="#FFF999" align="center" width="10%">노출기간</td>
    <td>
		<input type="text" id="sDt" name="StartDate" size="10" value="<%=sDt%>" />
		<input type="text" name="sTm" size="8" value="<%=sTm%>" /> ~
		<input type="text" id="eDt" name="EndDate" size="10" value="<%=eDt%>" />
		<input type="text" name="eTm" size="8" value="<%=eTm%>" />
    </td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center">고정 여부</td>
	<td><div style="float:left;"><input type="radio" name="topfixed" value="Y" <%=chkiif(topfixed = "Y","checked","")%> />사용함 &nbsp;&nbsp;&nbsp; <input type="radio" name="topfixed" value="N"  <%=chkiif(topfixed = "N" Or topfixed = "","checked","")%>/>사용안함</div> <div style="float:right;margin-top:5px;margin-right:10px;"></div></td>
</tr>
<tr bgcolor="#FFFFFF"><%	''2016-04-28 유태욱 추가(기획전1,마케팅2 구분) %>
	<td bgcolor="#FFF999" align="center">이벤트 구분</td>
	<td>
		<div style="float:left;">
			<input type="radio" name="evtgubun" value="1" <%=chkiif(evtgubun = "1" Or evtgubun = "","checked","")%> />기획전 이벤트 &nbsp;&nbsp;&nbsp; 
			<input type="radio" name="evtgubun" value="2"  <%=chkiif(evtgubun = "2","checked","")%>/>마케팅 이벤트
		</div>
		<div style="float:right;margin-top:5px;margin-right:10px;"></div>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center" height="40">구분</td>
	<td>
		<select name="gubun" onchange="onchgbox(this.value);" width="100" id="gubun">
			<option value="">=====구분선택=====</option>
			<option value="1" <%=chkiif(gubun="1","selected","")%>>Mobile & Apps</option>
			<option value="2" <%=chkiif(gubun="2","selected","")%>>Mobile</option>
			<option value="3" <%=chkiif(gubun="3","selected","")%>>Apps</option>
		</select>&nbsp;<br/>※ 구분 선택후 이벤트 코드를 입력 해주세요<br/>※ <span style="color:red">mobile & apps</span> 일때 동일한 이벤트 코드의 경우도 둘다 입력 해주세요
	</td>
</tr>
<tr bgcolor="#FFFFFF" id="sel22">
	<td bgcolor="#FFF999" align="center" >이벤트 이미지</td>
	<td width="45%">
		<input type="file" name="evtimg" class="file" title="이벤트 #1" require="N" style="width:80%;" />
		<% if mktimg<>"" then %>
		<br>
		<img src="<%= mktimg %>" width="300" /><br><%= mktimg %>
		<% end if %>
		<br/>
		※ 이벤트 이미지가 <span style="color:red">없을경우</span> 노출이 되지 않습니다.
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center" height="40">이미지 alt명</td>
	<td>
		<input type="text" name="altname" size="40" value="<%=altname%>"/>
	</td>
</tr>
<tr bgcolor="#FFFFFF" style="display:<%=chkiif(gubun=1 Or gubun=2,"block","none")%>" id="url1">
	<td bgcolor="#FFF999" align="center" height="40">모바일 이벤트 코드</td>
	<td>
		<input type="text" name="m_eventid" size="10" value="<%=m_eventid%>"/>
	</td>
</tr>
<tr bgcolor="#FFFFFF" style="display:<%=chkiif(gubun=1 Or gubun=3,"block","none")%>" id="url2">
	<td bgcolor="#FFF999" align="center" height="40">APP 이벤트 코드</td>
	<td>
		<input type="text" name="a_eventid" size="10" value="<%=a_eventid%>"/>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center">정렬 번호</td>
	<td><input type="text" name="sortnum" size="10" value="<%=chkiif(mode="modify",sortnum,"99")%>" maxlength="3"/></td>
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
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
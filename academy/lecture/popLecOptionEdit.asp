<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/academy/lib/academy_function.asp"-->
<!-- #include virtual="/academy/lib/classes/fingers_lecturecls.asp"-->
<%
dim mode
dim lec_idx, itemoption
dim oLect, oLectoption

Dim vOrderSerial
vOrderSerial = RequestCheckvar(Request("orderserial"),16)

lec_idx = RequestCheckvar(request("lec_idx"),10)
if lec_idx="" then lec_idx=0
mode= RequestCheckvar(request("mode"),16)
itemoption= RequestCheckvar(request("itemoption"),10)

dim sqlStr
dim ErrStr

'// 강좌 정보 확인
set oLect = new CLecture
oLect.FRectidx = lec_idx

if lec_idx<>"" then
	oLect.GetOneLecture
end if

if (oLect.FResultCount<1) then 
    response.write "권한이 없습니다."
    dbget.close()	:	response.End
end if

'// 강좌일정 접수
set oLectoption = new CLectOption
oLectoption.FRectidx = lec_idx
if lec_idx<>"" then
	oLectoption.GetLectOptionInfo
end if


dim i, j, k, pp
pp=0
%>
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script language='javascript'>
function EditOptionInfo(){
    var frm = document.frmEdit;
    
    if(frm.RegStartDate.length) {
	    for (var i=0;i<frm.RegStartDate.length;i++){
	        if (frm.RegStartDate[i].value.length<10){
	            alert('접수 시작일을 입력하세요.');
	            frm.RegStartDate[i].focus();
	            return;
	        }
	
	        if (frm.RegEndDate[i].value.length<10){
	            alert('접수 종료을 입력하세요.');
	            frm.RegEndDate[i].focus();
	            return;
	        }
	
	        if (frm.LecSDay[i].value.length<10){
	            alert('강좌 시작일을 입력하세요.');
	            frm.LecSDay[i].focus();
	            return;
	        }
	        if (frm.LecSTime[i].value.length<5){
	            alert('강좌 시작시간을 입력하세요.');
	            frm.LecSTime[i].focus();
	            return;
	        }
	
	        if (frm.LecETime[i].value.length<5){
	            alert('강좌 종료시간을 입력하세요.');
	            frm.LecETime[i].focus();
	            return;
	        }
	    }
    }
    else {
        if (frm.RegStartDate.value.length<10){
            alert('접수 시작일을 입력하세요.');
            frm.RegStartDate.focus();
            return;
        }

        if (frm.RegEndDate.value.length<10){
            alert('접수 종료을 입력하세요.');
            frm.RegEndDate.focus();
            return;
        }

        if (frm.LecSDay.value.length<10){
            alert('강좌 시작일을 입력하세요.');
            frm.LecSDay.focus();
            return;
        }
        if (frm.LecSTime.value.length<5){
            alert('강좌 시작시간을 입력하세요.');
            frm.LecSTime.focus();
            return;
        }

        if (frm.LecETime.value.length<5){
            alert('강좌 종료시간을 입력하세요.');
            frm.LecETime.focus();
            return;
        }
    }
   
    if (confirm('수정 하시겠습니까?')){
        frm.mode.value="modi";
        frm.submit();
    }
}


function NewSaveOption(){
	var frm = document.frmEdit;
    if (frm.tmpRegSDt.value.length<10){
        alert('접수 시작일을 입력하세요.');
        frm.tmpRegSDt.focus();
        return;
    }
    if (frm.tmpRegEDt.value.length<10){
        alert('접수 종료일을 입력하세요.');
        frm.tmpRegEDt.focus();
        return;
    }

    if (frm.tmpLecSDt.value.length<10){
        alert('강의 시작일을 입력하세요.');
        frm.tmpLecSDt.focus();
        return;
    }
    if (frm.tmpLecStime.value.length<5){
        alert('강의 시작시간을 입력하세요.');
        frm.tmpLecStime.focus();
        return;
    }

    if (frm.tmpLecEtime.value.length<5){
        alert('강의 종료시간을 입력하세요.');
        frm.tmpLecEtime.focus();
        return;
    }

    if (confirm('저장 하시겠습니까?')){
        frm.mode.value="write";
        frm.submit();
    }
}

// ============================================================================


function AddOptionLyrSw(){
    var fm = document.all;
    if(fm.AddOptForm.style.display=="") {
    	fm.OptAddBtn.value="일정추가 ▼";
    	fm.OptAddBtn.blur();
    	fm.AddOptForm.style.display = "none";
    } else {
    	fm.OptAddBtn.value="일정추가 ▲";
    	fm.OptAddBtn.blur();
    	fm.AddOptForm.style.display = "";
    }
}

function chgUsing(uVl,uNo) {
	var fm = document.frmEdit;
	if(fm.isUsing.length) {
		fm.isUsing[uNo].value=uVl;
	} else {
		fm.isUsing.value=uVl;
	}
}

//옵션변경
function changethisoption(option)
{
	document.optionchange.option.value = option;
	document.optionchange.submit();
}
</script>

<form name="optionchange" action="/cscenterv2/lecture/lecture_option_change_proc.asp" method="post" style="margin:0px;">
<input type="hidden" name="lec_idx" value="<%= lec_idx %>">
<input type="hidden" name="orderserial" value="<%= vOrderSerial %>">
<input type="hidden" name="option" value="">
</form>

<!-- 표 상단바 시작-->
<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a" bgcolor="#999999">
<tr height="25" valign="bottom" bgcolor="F4F4F4">
    <td valign="top" bgcolor="F4F4F4">
    	<b>강의 일정(옵션) 수정</b><br>

    	<br>- 일정을 추가 또는 삭제할수 있습니다.
    	<br>- 접수 내역이 있는 일정은 삭제가 불가능합니다.(사용안함 으로 수정하세요)
    </td>
</tr>
</table>
<p>
<!-- 표 상단바 끝-->
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor=#BABABA>
<tr>
	<td width=80 height="25" bgcolor="#DDDDFF" align="center">강의코드</td>
	<td  bgcolor="#FFFFFF"><%= lec_idx %></td>
	<td width=240 bgcolor="#DDDDFF" align="center">일정 선택 미리보기</td>
</tr>
<tr>
	<td width=80 height="25" bgcolor="#DDDDFF" align="center">강의명</td>
	<td bgcolor="#FFFFFF"><%= oLect.FOneItem.Flec_title %></td>
	<td width=200 bgcolor="#FFFFFF" rowspan="2" align="center">
	<%= getLecOptionBoxHTML(lec_idx,"lecOption","") %>
	</td>
</tr>
<tr>
	<td width=80 height="25" bgcolor="#DDDDFF" align="center">강사</td>
	<td bgcolor="#FFFFFF"><%= oLect.FOneItem.Flecturer_name %> (<%= oLect.FOneItem.Flecturer_id %>)</td>
</tr>
</table>

<p>

<table width="100%" align="center" cellpadding="3" cellspacing="1" border="0" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmEdit" method="post" action="do_adminLecOptionEdit.asp">
<input type="hidden" name="lec_idx" value="<%= lec_idx %>">
<input type="hidden" name="orderserial" value="<%= vOrderSerial %>">
<input type="hidden" name="mode" value="">
	<tr height="25" bgcolor="FFFFFF">
	    
		<td colspan="8"> 
		    <table width="100%" cellpadding="0" cellspacing="0" border="0" class="a" >
		    <tr>
		        <td>등록된 일정 리스트</td>
		        <td width="80" align="right"><input name="OptAddBtn" id="OptAddBtn" type="button" class="button" value="일정추가 ▼" onClick="AddOptionLyrSw();"></td>
		    </tr>
		    </table>
		</td>
	</tr>
	<tr id="AddOptForm" height="25" bgcolor="FFFFFF" style="display:none">
		<td colspan="8"> 
		    <table width="100%" cellpadding="2" cellspacing="1" border="0" class="a" style="border:1px solid #c0c0c0;">
		    <tr><td colspan="2" bgcolor="#DDDDFF">강의일정 입력</td></tr>
		    <tr>
		        <td width="100" bgcolor="EDEDFF" align="center">접수기간</td>
		        <td>
		        	시작일 <input id="tmpRegSDt" name="tmpRegSDt" class="text" size="10" maxlength="10" /> <img src="http://scm.10x10.co.kr/images/calicon.gif" id="tmpRegSDt_trigger" border="0" style="cursor:pointer" align="absmiddle" /> ~
		        	종료일 <input id="tmpRegEDt" name="tmpRegEDt" class="text" size="10" maxlength="10" /> <img src="http://scm.10x10.co.kr/images/calicon.gif" id="tmpRegEDt_trigger" border="0" style="cursor:pointer" align="absmiddle" />
		        </td>
		    </tr>
		    <tr>
		        <td width="100" bgcolor="EDEDFF" align="center">일정명</td>
		        <td><input type="text" name="tmpLecOptName" size="64" maxlength="128"></td>
		    </tr>
		    <tr>
		        <td width="100" bgcolor="EDEDFF" align="center">강의일자</td>
		        <td>
					날짜 <input id="tmpLecSDt" name="tmpLecSDt" class="text" size="10" maxlength="10" /> <img src="http://scm.10x10.co.kr/images/calicon.gif" id="tmpLecSDt_trigger" border="0" style="cursor:pointer" align="absmiddle" />&nbsp;/
					시간 <input type="text" name="tmpLecStime" size="5" maxlength="5" class="text" value="00:00" onfocus="this.style.backgroundColor='FFE0E0'" onblur="this.style.backgroundColor='FFFFFF'">
					~
					<input type="text" name="tmpLecEtime" size="5" maxlength="5" class="text" value="00:00" onfocus="this.style.backgroundColor='FFE0E0'" onblur="this.style.backgroundColor='FFFFFF'">
					<script language="javascript">
						var CAL_Start = new Calendar({
							inputField : "tmpRegSDt", trigger    : "tmpRegSDt_trigger",
							onSelect: function() {
								var date = Calendar.intToDate(this.selection.get());
								CAL_End.args.min = date;
								CAL_End.redraw();
								this.hide();
							}, bottomBar: true, dateFormat: "%Y-%m-%d"
						});
						var CAL_End = new Calendar({
							inputField : "tmpRegEDt", trigger    : "tmpRegEDt_trigger",
							onSelect: function() {
								var date = Calendar.intToDate(this.selection.get());
								CAL_Start.args.max = date;
								CAL_Start.redraw();
								this.hide();
							}, bottomBar: true, dateFormat: "%Y-%m-%d"
						});
						var CAL_Prize = new Calendar({
							inputField : "tmpLecSDt", trigger    : "tmpLecSDt_trigger",
							onSelect: function() { this.hide(); }, bottomBar: true, dateFormat: "%Y-%m-%d"
						});
					</script>
		        </td>
		    </tr>
		    <tr><td colspan="2" bgcolor="#F0F0F0" align="right" style="padding-right:5px"><input type="button" value="일정저장" class="button" onClick="NewSaveOption()"></td></tr>
		    </table>
		</td>
	</tr>
	<% if oLectoption.FResultCount<1 then %>
    <tr height="25" bgcolor="#FFFFFF">
	    <td colspan="8" align=center>등록된 일정이 없습니다.</td>
    </tr>
    <% else %>
	    <!-- 단일옵션  -->
	    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
        	<td width="40">코드</td>
        	<td width="120">접수기간</td>
        	<td >강좌일시</td>
        	<td width="60">사용<br>여부</td>
        	<td width="40">정원</td>
        	<td width="40">신청<br>인원</td>
        	<td width="40">남은<br>인원</td>
        	<td width="40">마감<br>여부</td>
        </tr>
       	<% for k=0 to oLectoption.FResultCount -1 %>
        <tr align="center" bgcolor="<%= ChkIIF(oLectoption.FItemList(k).Fisusing="Y","#FFFFFF","#DDDDDD") %>">
        	<td><%= oLectoption.FItemList(k).FlecOption %><input type="hidden" name="lecOption" value="<%= oLectoption.FItemList(k).FlecOption %>">
        		<% If vOrderSerial <> "" Then %><br><b><font color="blue"><a href="javascript:changethisoption('<%= oLectoption.FItemList(k).FlecOption %>');">[선택]</a></font></b><% End If %>
        	</td>
        	<td>
        		시작: <input type="text" class="text" name="RegStartDate" value="<%= FormatDateTime(oLectoption.FItemList(k).FRegStartDate,2) %>" size="10" maxlength="10" onfocus="this.style.backgroundColor='FFE0E0'" onblur="this.style.backgroundColor='FFFFFF'"><br>
        		종료: <input type="text" class="text" name="RegEndDate" value="<%= FormatDateTime(oLectoption.FItemList(k).FRegEndDate,2) %>" size="10" maxlength="10" onfocus="this.style.backgroundColor='FFE0E0'" onblur="this.style.backgroundColor='FFFFFF'">
        	</td>
        	<td align="left">
        		<input type="text" name="LecOptionName" size="30" maxlength="128" value="<%=oLectoption.FItemList(k).FlecOptionName%>"><br>
        		<input type="text" class="text" name="LecSDay" value="<%= FormatDateTime(oLectoption.FItemList(k).FlecStartDate,2) %>" size="10" maxlength="10" onfocus="this.style.backgroundColor='FFE0E0'" onblur="this.style.backgroundColor='FFFFFF'">
        		<input type="text" class="text" name="LecSTime" value="<%= FormatDateTime(oLectoption.FItemList(k).FlecStartDate,4) %>" size="4" maxlength="5" onfocus="this.style.backgroundColor='FFE0E0'" onblur="this.style.backgroundColor='FFFFFF'">~
        		<input type="text" class="text" name="LecETime" value="<%= FormatDateTime(oLectoption.FItemList(k).FlecEndDate,4) %>" size="4" maxlength="5" onfocus="this.style.backgroundColor='FFE0E0'" onblur="this.style.backgroundColor='FFFFFF'">
        	</td>
        	<td>
        		<input type="radio" name="tmpUse<%=k%>" value="Y" <%= ChkIIF(oLectoption.FItemList(k).Fisusing="Y","checked","") %> onClick="chgUsing(this.value,<%=k%>)">Y
        		<input type="radio" name="tmpUse<%=k%>" value="N" <%= ChkIIF(oLectoption.FItemList(k).Fisusing="N","checked","") %> onClick="chgUsing(this.value,<%=k%>)">N
        		<input type="hidden" name="isUsing" value="<%=oLectoption.FItemList(k).Fisusing%>">
        	</td>
        	<td><input type="text" name="limit_count" size="1" class="text" value="<%= oLectoption.FItemList(k).Flimit_count %>" onfocus="this.style.backgroundColor='FFE0E0'" onblur="this.style.backgroundColor='FFFFFF'"></td>
        	<td><input type="text" name="limit_sold" size="1" class="text" value="<%= oLectoption.FItemList(k).Flimit_sold %>" onfocus="this.style.backgroundColor='FFE0E0'" onblur="this.style.backgroundColor='FFFFFF'"></td>
        	<td><%= oLectoption.FItemList(k).Flimit_count-oLectoption.FItemList(k).Flimit_sold %></td>
        	<td><% if oLectoption.FItemList(k).IsOptionSoldOut then %><font color="red">마감</font><% end if %></td>
            <% pp = pp + 1 %>
        </tr>
       	<% next %>
        </tr>
	<% end if %>
</form>
</table>

<p>
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor=#FFFFFF>
<tr height="30">
    <td align="center"><input type="button" value="일정 내용 수정" onClick="EditOptionInfo();"></td>
</tr>
</table>
<%
set oLect = Nothing
set oLectoption = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
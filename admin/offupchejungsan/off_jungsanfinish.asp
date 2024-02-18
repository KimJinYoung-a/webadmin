<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/offshopclass/offjungsancls.asp"-->

<%
dim research, segumtype
dim thismonth

research = request("research")
segumtype = request("segumtype")


thismonth = Left(CStr(DateSerial(year(now()),month(now())-1,1)),7)
%>


<script language='javascript'>


function PopJungsanUpload(){
	var popwin = window.open("pop_off_jungsan_upload.asp","pop_off_jungsan_upload","width=800 height=800 scrollbars=yes resizable=yes");
	popwin.focus();
}

</script>


<!-- 표 상단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
   	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="research" value="on">
   	<tr height="10" valign="bottom">
        <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_02.gif"></td>
        <td background="/images/tbl_blue_round_02.gif"></td>
        <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<tr height="25" valign="top">
        <td background="/images/tbl_blue_round_04.gif"></td>
        <td valign="top">
        	<input type="button" value="정산업로드파일" onclick="PopJungsanUpload();">
        </td>
        <td valign="top" align="right">
        	<a href="javascript:document.frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
        </td>
        <td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	</form>
</table>
<!-- 표 상단바 끝-->


<%
dim ipkumregdate
ipkumregdate = request("ipkumregdate")


dim ooffjungsan
set ooffjungsan = new COffJungsan
ooffjungsan.FRectNotIncludeWonChon = "on"
ooffjungsan.FRectYYYYMM = thismonth
ooffjungsan.FRectbankingupflag = "Y"

ooffjungsan.JungsanFixedList

dim ipsum,i
ipsum =0
%>
<script language='javascript'>
function ipkumfinish(frm,iidx){
	if (frm.ipkumregdate.value.length<1){
		alert('입금일을 입력하세요.');
		frm.ipkumregdate.focus();
		return;
	}

	frm.idx.value= iidx;

	var ret = confirm('진행하시겠습니까?');

	if (ret){
		var popwin = window.open("","regipkumfinish","width=300 height=300");
		popwin.focus();
		frm.target = "regipkumfinish";
		frm.submit();
	}
}

function delbankingup(iidx){
	var ret = confirm('삭제 하시겠습니까?');

	if (ret){
		var popwin = window.open("bankingupflag_process.asp?mode=delflag&id=" + iidx,"regipkumfinish","width=100 height=100");
		popwin.focus();
	}
}

function batchipkumfinish(frm){
	if (frmip.ipkumregdate.value.length<1){
		alert('입금일을 입력하세요.');
		calendarOpen(frmip.ipkumregdate);
		return;
	}


	if (confirm(frmip.ipkumregdate.value + '로 입금확인 진행 하시겠습니까?')){
		frm.ipkumregdate.value=frmip.ipkumregdate.value;
		frm.submit();
	}
}

</script>

<!-- 표 중간바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
	<form name=frmip method=post action="">
    <input type=hidden name=rd_state value=7>
    <input type="hidden" name="mode" value="ipkumfinish">
    <input type="hidden" name="idx" value="">
	<tr>
		<td height="1" colspan="15" bgcolor="<%= adminColor("tablebg") %>"></td>
	</tr>
    <tr height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td align="left">
	        입금일 : <input type=text name=ipkumregdate value="<%= ipkumregdate %>" size=10 maxlength=10 readonly >
	    	<a href="javascript:calendarOpen(frmip.ipkumregdate);"><img src="/images/calicon.gif" border="0" align="absmiddle" height=21></a>
	    	(2004-06-30)
    		<input type="button" value="전체입금완료진행" onclick="batchipkumfinish(frmbatch);">
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    </form>
</table>
<!-- 표 중간바 끝-->





<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
    <tr bgcolor="#FFFFFF">
    	<td colspan="15" >금월(<%= thismonth %>) 세금계산서 (<%= ooffjungsan.FresultCount %>건)</td>
    </tr>
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="60">정산월</td>
		<td width="70">발행일</td>
		<td width="40">정산일</td> 
		<td width="120">브랜드ID</td>
      	<td width="150">예금주</td>
		<td width="60">상태</td>
		<td width="60">은행</td>
		<td width="80">계좌</td>
		<td width="80">정산금액</td>
		<td>업체명</td>
		<td width="30">삭제</td>
	</tr>
<form name="frmbatch" method="post" action="bankingupflag_process.asp">
<input type="hidden" name="mode" value="ipkumfinish">
<input type="hidden" name="ipkumregdate" value="">
<% for i=0 to ooffjungsan.FresultCount-1 %>
<%
ipsum = ipsum + ooffjungsan.FItemList(i).Ftot_jungsanprice
%>
	<input type=hidden name="checkone" value="<%= ooffjungsan.FItemList(i).FIdx %>">
	<% if ooffjungsan.FItemList(i).Ftot_jungsanprice<0 then %>
	<tr align="center" bgcolor="<%= adminColor("dgray") %>">
	<% else %>
	<tr align="center" bgcolor="#FFFFFF">
	<% end if %>
		<td><%= ooffjungsan.FItemList(i).Fyyyymm %></td>
		<td>
			<% if Left(ooffjungsan.FItemList(i).Ftaxregdate,7) = Left(CStr(now()),7) then %>
			<font color="red"><%= ooffjungsan.FItemList(i).Ftaxregdate %></font>
			<% else %>
			<font color="blue"><%= ooffjungsan.FItemList(i).Ftaxregdate %></font>
			<% end if %>
		</td>
		<td><%= ooffjungsan.FItemList(i).Fjungsan_date_off %></td>
		<td><a href="javascript:PopUpcheBrandInfoEdit('<%= ooffjungsan.FItemList(i).Fmakerid %>')"><%= ooffjungsan.FItemList(i).Fmakerid %></a></td>
		<td><%= ooffjungsan.FItemList(i).Fjungsan_acctname %></td>
		<td><font color="<%= ooffjungsan.FItemList(i).GetStateColor %>"><%= ooffjungsan.FItemList(i).GetStateName %></font></td>
		<td><%= ooffjungsan.FItemList(i).Fipkum_bank %></td>
		<td><%= ooffjungsan.FItemList(i).Fipkum_acctno %></td>
		<td align="right"><%= FormatNumber(ooffjungsan.FItemList(i).Ftot_jungsanprice,0) %></td>
		<td><%= ooffjungsan.FItemList(i).Fcompany_name %></td>
		<td>
		<a href="javascript:delbankingup('<%= ooffjungsan.FItemList(i).Fidx %>')">
		x
		</a>
		</td>
	</tr>
<% next %>
	<tr bgcolor="#FFFFFF">
		<td colspan="8"></td>
		<td align="right"><%= FormatNumber(ipsum,0) %></td>
		<td colspan="2"></td>
	</tr>
</table>

<%
ooffjungsan.FRectYYYYMM = ""
ooffjungsan.FRectNotIncludeWonChon = "on"
ooffjungsan.FRectNotYYYYMM = thismonth
ooffjungsan.FRectbankingupflag = "Y"

ooffjungsan.JungsanFixedList



ipsum =0
%>

<br>
<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
    <tr bgcolor="#FFFFFF">
    	<td colspan="15" >전월 세금계산서 (<%= ooffjungsan.FresultCount %>건)</td>
    </tr>
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="60">정산월</td>
		<td width="70">발행일</td>
		<td width="40">정산일</td> 
		<td width="120">브랜드ID</td>
      	<td width="150">예금주</td>
		<td width="60">상태</td>
		<td width="60">은행</td>
		<td width="80">계좌</td>
		<td width="80">정산금액</td>
		<td>업체명</td>
		<td width="30">삭제</td>
     </tr>
<% for i=0 to ooffjungsan.FresultCount-1 %>
<%
ipsum = ipsum + ooffjungsan.FItemList(i).Ftot_jungsanprice
%>
	<input type=hidden name="checkone" value="<%= ooffjungsan.FItemList(i).FIdx %>">
	<% if ooffjungsan.FItemList(i).Ftot_jungsanprice<0 then %>
	<tr align="center" bgcolor="<%= adminColor("dgray") %>">
	<% else %>
	<tr align="center" bgcolor="#FFFFFF">
	<% end if %>
		<td><%= ooffjungsan.FItemList(i).Fyyyymm %></td>
		<td>
			<% if Left(ooffjungsan.FItemList(i).Ftaxregdate,7) = Left(CStr(now()),7) then %>
			<font color="red"><%= ooffjungsan.FItemList(i).Ftaxregdate %></font>
			<% else %>
			<font color="blue"><%= ooffjungsan.FItemList(i).Ftaxregdate %></font>
			<% end if %>
		</td>
		<td><%= ooffjungsan.FItemList(i).Fjungsan_date_off %></td>
		<td><a href="javascript:PopUpcheBrandInfoEdit('<%= ooffjungsan.FItemList(i).Fmakerid %>')"><%= ooffjungsan.FItemList(i).Fmakerid %></a></td>
		<td><%= ooffjungsan.FItemList(i).Fjungsan_acctname %></td>
		<td><font color="<%= ooffjungsan.FItemList(i).GetStateColor %>"><%= ooffjungsan.FItemList(i).GetStateName %></font></td>
		<td><%= ooffjungsan.FItemList(i).Fipkum_bank %></td>
		<td><%= ooffjungsan.FItemList(i).Fipkum_acctno %></td>
		<td align="right"><%= FormatNumber(ooffjungsan.FItemList(i).Ftot_jungsanprice,0) %></td>
		<td><%= ooffjungsan.FItemList(i).Fcompany_name %></td>
		<td>
		<a href="javascript:delbankingup('<%= ooffjungsan.FItemList(i).Fidx %>')">
		x
		</a>
		</td>
	</tr>
<% next %>
	<tr bgcolor="#FFFFFF">
		<td colspan="8"></td>
		<td align="right"><%= FormatNumber(ipsum,0) %></td>
		<td colspan="2"></td>
	</tr>
</table>

<%
ooffjungsan.FRectYYYYMM = ""
ooffjungsan.FRectNotYYYYMM = ""
ooffjungsan.FRectNotIncludeWonChon = ""
ooffjungsan.FRectOnlyIncludeWonChon = "on"
ooffjungsan.FRectbankingupflag = "Y"

ooffjungsan.JungsanFixedList

ipsum =0
%>
<br>

<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
    <tr bgcolor="#FFFFFF">
    	<td colspan="15" >원천징수 대상자 (<%= ooffjungsan.FresultCount %>건)</td>
    </tr>
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="60">정산월</td>
		<td width="70">발행일</td>      																																																																																																
		<td width="40">정산일</td>        																																																																																																
		<td width="120">브랜드ID</td>
      	<td width="100">예금주</td>																																																																							
		<td width="60">상태</td>          																																																																																																
		<td width="60">은행</td>          																																																																																																
		<td width="80">계좌</td>          																																																																																																
		<td width="60">확정금액</td>      																																																																																																																
		<td width="60">정산금액*0.967</td>																																																																																																																
		<td>업체명</td>                   																																																																																																																
		<td width="30">삭제</td> 																																																																																																																
     </tr>
<% for i=0 to ooffjungsan.FresultCount-1 %>
<%
ipsum = ipsum + fix(ooffjungsan.FItemList(i).Ftot_jungsanprice*0.967)
%>
	<input type=hidden name="checkone" value="<%= ooffjungsan.FItemList(i).FIdx %>">
	<% if ooffjungsan.FItemList(i).Ftot_jungsanprice<0 then %>
	<tr align="center" bgcolor="<%= adminColor("dgray") %>">
	<% else %>
	<tr align="center" bgcolor="#FFFFFF">
	<% end if %>
		<td><%= ooffjungsan.FItemList(i).Fyyyymm %></td>
		<td>
			<% if Left(ooffjungsan.FItemList(i).Ftaxregdate,7) = Left(CStr(now()),7) then %>
			<font color="red"><%= ooffjungsan.FItemList(i).Ftaxregdate %></font>
			<% else %>
			<font color="blue"><%= ooffjungsan.FItemList(i).Ftaxregdate %></font>
			<% end if %>
		</td>
		<td><%= ooffjungsan.FItemList(i).Fjungsan_date_off %></td>
		<td><a href="javascript:PopUpcheBrandInfoEdit('<%= ooffjungsan.FItemList(i).Fmakerid %>')"><%= ooffjungsan.FItemList(i).Fmakerid %></a></td>
		<td><%= ooffjungsan.FItemList(i).Fjungsan_acctname %></td>
		<td><font color="<%= ooffjungsan.FItemList(i).GetStateColor %>"><%= ooffjungsan.FItemList(i).GetStateName %></font></td>
		<td><%= ooffjungsan.FItemList(i).Fipkum_bank %></td>
		<td><%= ooffjungsan.FItemList(i).Fipkum_acctno %></td>
		<td align="right"><%= FormatNumber(ooffjungsan.FItemList(i).Ftot_jungsanprice,0) %></td>
		<td align="right"><%= FormatNumber(fix(ooffjungsan.FItemList(i).Ftot_jungsanprice*0.967),0) %></td>
		<td><%= ooffjungsan.FItemList(i).Fcompany_name %></td>
		<td>
		<a href="javascript:delbankingup('<%= ooffjungsan.FItemList(i).Fidx %>')">
		x
		</a>
		</td>
	</tr>
<% next %>
	<tr bgcolor="#FFFFFF">
		<td colspan="9"></td>
		<td align="right"><%= FormatNumber(ipsum,0) %></td>
		<td colspan="1"></td>
		<td colspan="1"></td>
	</tr>
</table>

<%
set ooffjungsan = Nothing
%>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->

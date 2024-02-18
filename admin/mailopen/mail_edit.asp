<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  메일진 통계
' History : 2008.01.21 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/report/reportcls.asp"-->

<%
Dim omd ,idx,mode
	idx = request("idx")
	mode = request("mode")
	
	If idx = "" Then idx=0

set omd = New CMailzineOne
	omd.GetMailingOne idx
%>

<script language="JavaScript">

function TnMailDataReg(){
	if(frm.title.value == ""){
		alert("발송이름을 적어주세요");
		frm.title.focus();
	}
	else if(frm.gubun.value == ""){
		alert("발송구분을 적어주세요");
		frm.gubun.focus();
	}
	else if(frm.startdate.value == ""){
		alert("발송시작시간을 적어주세요");
		frm.startdate.focus();
	}
	else if(frm.enddate.value == ""){
		alert("발송종료시간을 적어주세요");
		frm.enddate.focus();
	}
	else if(frm.reenddate.value == ""){
		alert("재발송종료시간을 적어주세요");
		frm.reenddate.focus();
	}
	else if(frm.totalcnt.value == ""){
		alert("총대상자수를 적어주세요");
		frm.totalcnt.focus();
	}
	else if(frm.realcnt.value == ""){
		alert("실발송통수를 적어주세요");
		frm.realcnt.focus();
	}
	else if(frm.realpct.value == ""){
		alert("실발송비율을 적어주세요");
		frm.realpct.focus();
	}
	else if(frm.filteringcnt.value == ""){
		alert("필터링통수를 적어주세요");
		frm.filteringcnt.focus();
	}
	else if(frm.filteringpct.value == ""){
		alert("필터링비율을 적어주세요");
		frm.filteringpct.focus();
	}
	else if(frm.successcnt.value == ""){
		alert("성공발송통수를 적어주세요");
		frm.successcnt.focus();
	}
	else if(frm.successpct.value == ""){
		alert("성공율을 적어주세요");
		frm.successpct.focus();
	}
	else if(frm.failcnt.value == ""){
		alert("실패발송통수를 적어주세요");
		frm.failcnt.focus();
	}
	else if(frm.failpct.value == ""){
		alert("실패율을 적어주세요");
		frm.failpct.focus();
	}
	else if(frm.opencnt.value == ""){
		alert("오픈통수를 적어주세요");
		frm.opencnt.focus();
	}
	else if(frm.openpct.value == ""){
		alert("오픈율을 적어주세요");
		frm.openpct.focus();
	}
	else if(frm.noopencnt.value == ""){
		alert("미오픈통수를 적어주세요");
		frm.noopencnt.focus();
	}
	else if(frm.noopenpct.value == ""){
		alert("미오픈율을 적어주세요");
		frm.noopenpct.focus();
	}
	else if(frm.isusing.value == ""){
		alert("사용여부를 선택해 주세요");
		frm.isusing.focus();		
	}
	else if(frm.mailergubun.value == ""){
		alert("발송메일러를 선택해 주세요");
		frm.mailergubun.focus();		
	}	
	else{
		frm.submit();
	}
}

</script>

<table width="100%" border="0" class="a" cellpadding="1" cellspacing="1" bgcolor="#BABABA" align="center">
<form action="/admin/mailopen/mail_process.asp" name="frm" method="post">
<input type="hidden" name="mode" value="<% = mode %>">
<input type="hidden" name="idx" value="<% = idx %>">
<tr bgcolor=#FFFFFF>
	<td align="center">발송구분</td>
	<td align="left">
		<input name="gubun" size="25" type="text" value="<% = omd.fgubun %>">
	</td>
</tr>
<tr bgcolor=#FFFFFF>
	<td align="center">발송이름</td>
	<td align="left"><input name="title" size="25" type="text" value="<% = omd.Ftitle %>"></td>
</tr>
<tr bgcolor=#FFFFFF>
	<td align="center">발송시작시간</td>
	<td align="left"><input name="startdate" size="25" type="text" value="<% = omd.Fstartdate %>"></td>
</tr>	
<tr bgcolor=#FFFFFF>
	<td align="center">발송종료시간</td>
	<td align="left"><input name="enddate" size="25" type="text" value="<% = omd.Fenddate %>"></td>
</tr>			
<tr bgcolor=#FFFFFF>
	<td align="center">재발송종료시간</td>
	<td align="left"><input name="reenddate" size="25" type="text" value="<% = omd.Freenddate %>"></td>
</tr>	
<tr bgcolor=#FFFFFF>
	<td align="center">총대상자수</td>
	<td align="left"><input name="totalcnt" size="15" type="text" value="<% = omd.Ftotalcnt %>"></td>
</tr>			
<tr bgcolor=#FFFFFF>
	<td align="center">실발송통수(비율)</td>
	<td align="left"><input name="realcnt" size="15" type="text" value="<% = omd.Frealcnt %>">
	<input name="realpct" size="10" type="text" value="<% = omd.Frealpct %>">
	</td>
</tr>	
<tr bgcolor=#FFFFFF>
	<td align="center">필터링통수</td>
	<td align="left"><input name="filteringcnt" size="15" type="text" value="<% = omd.Ffilteringcnt %>">
	<input name="filteringpct" size="10" type="text" value="<% = omd.Ffilteringpct %>">
	</td>
</tr>	
<tr bgcolor=#FFFFFF>
	<td align="center">성공발송통수(비율)</td>
	<td align="left"><input name="successcnt" size="15" type="text" value="<% = omd.Fsuccesscnt %>">
	<input name="successpct" size="10" type="text" value="<% = omd.Fsuccesspct %>">
	</td>
</tr>		
<tr bgcolor=#FFFFFF>
	<td align="center">실패발송통수</td>
	<td align="left"><input name="failcnt" size="15" type="text" value="<% = omd.Ffailcnt %>">
	<input name="failpct" size="10" type="text" value="<% = omd.Ffailpct %>">
	</td>
</tr>		
<tr bgcolor=#FFFFFF>
	<td align="center">오픈통수</td>
	<td align="left"><input name="opencnt" size="15" type="text" value="<% = omd.Fopencnt %>">
	<input name="openpct" size="10" type="text" value="<% = omd.Fopenpct %>">
	</td>
</tr>			
<tr bgcolor=#FFFFFF>
	<td align="center">미오픈통수</td>
	<td align="left"><input name="noopencnt" size="15" type="text" value="<% = omd.Fnoopencnt %>">
	<input name="noopenpct" size="10" type="text" value="<% = omd.Fnoopenpct %>">
	</td>
</tr>
<tr bgcolor=#FFFFFF>
	<td align="center">사용여부</td>
	<td align="left">
		<select name="isusing">
			<option value="Y" <% if omd.Fisusing = "Y" then response.write " selected" %>>Y</option>
			<option value="N" <% if omd.Fisusing = "N" then response.write " selected" %>>N</option>
		</select>
	</td>
</tr>
<tr bgcolor=#FFFFFF>
	<td align="center">발송메일러</td>
	<td align="left">
		<% drawmailergubun "mailergubun" , omd.fmailergubun , "" %>
	</td>
</tr>
<tr bgcolor=#FFFFFF>
	<td align="center" colspan=2>
		<input type="button" value="저장" onclick="javascript:TnMailDataReg();" class="button">
	</td>	
</tr>
</form>
</table>

<% set omd = nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
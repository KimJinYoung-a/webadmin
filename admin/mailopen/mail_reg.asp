<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  메일진 통계
' History : 2008.01.21 한용민 개발
'			2012.05.09 김진영 구분 추가
'			2012.12.04 김진영 오프라인항목 추가
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<%
' 복사내역에서 필요한 부분만 추출해 낸다. ###############################################
Dim ret_txt,ret_txt_re , i
	ret_txt = request("ret_txt")
	ret_txt  = replace(ret_txt,vbcrlf," bbb ")
	ret_txt_re = split(ret_txt," bbb ")

Dim ret_txt_seach,ret_txt_seach_re
	ret_txt_seach = ""
	If ret_txt <> "" Then
		ret_txt_seach = ret_txt_seach + ret_txt_re(1)&" bbb "
		ret_txt_seach = ret_txt_seach + ret_txt_re(8)&" bbb "
		ret_txt_seach = ret_txt_seach + ret_txt_re(9)&" bbb "
		ret_txt_seach = ret_txt_seach + ret_txt_re(10)&" bbb "
		ret_txt_seach = ret_txt_seach + ret_txt_re(11)&" bbb "
		ret_txt_seach = ret_txt_seach + ret_txt_re(13)&" bbb "
		ret_txt_seach_re = split(ret_txt_seach," ")

		If ret_txt_seach_re(0) <> "mailzine" and ret_txt_seach_re(0) <> "mailzine_비회원" and ret_txt_seach_re(0) <> "mailzine_fingers" and ret_txt_seach_re(0) <> "mailzine_핑거스비회원" and ret_txt_seach_re(0) <> "OFFLINE" Then
			response.write "<script>" &_
						   "	alert('캠페인제목을 확인하세요');" &_
						   "	location.replace('/admin/mailopen/mail_reg.asp?mode=add');" &_
						   "</script>" &_
			response.end
		End If
	End If
' 복사내역에서 필요한 부분만 추출해 낸다. ###############################################
%>
<script language="JavaScript">
function TnMailDataReg(){
	if(frm.title.value == ""){
		alert("발송이름을 적어주세요");
		frm.title.focus();
	}else if(frm.gubun.value == ""){
		alert("발송구분을 적어주세요");
		frm.gubun.focus();
	}else if(frm.startdate.value == ""){
		alert("발송시작시간을 적어주세요");
		frm.startdate.focus();
	}else if(frm.enddate.value == ""){
		alert("발송종료시간을 적어주세요");
		frm.enddate.focus();
	}else if(frm.reenddate.value == ""){
		alert("재발송종료시간을 적어주세요");
		frm.reenddate.focus();
	}else if(frm.totalcnt.value == ""){
		alert("총대상자수를 적어주세요");
		frm.totalcnt.focus();
	}else if(frm.realcnt.value == ""){
		alert("실발송통수를 적어주세요");
		frm.realcnt.focus();
	}else if(frm.realpct.value == ""){
		alert("실발송비율을 적어주세요");
		frm.realpct.focus();
	}else if(frm.filteringcnt.value == ""){
		alert("필터링통수를 적어주세요");
		frm.filteringcnt.focus();
	}else if(frm.filteringpct.value == ""){
		alert("필터링비율을 적어주세요");
		frm.filteringpct.focus();
	}else if(frm.successcnt.value == ""){
		alert("성공발송통수를 적어주세요");
		frm.successcnt.focus();
	}else if(frm.successpct.value == ""){
		alert("성공율을 적어주세요");
		frm.successpct.focus();
	}else if(frm.failcnt.value == ""){
		alert("실패발송통수를 적어주세요");
		frm.failcnt.focus();
	}else if(frm.failpct.value == ""){
		alert("실패율을 적어주세요");
		frm.failpct.focus();
	}else if(frm.opencnt.value == ""){
		alert("오픈통수를 적어주세요");
		frm.opencnt.focus();
	}else if(frm.openpct.value == ""){
		alert("오픈율을 적어주세요");
		frm.openpct.focus();
	}else if(frm.noopencnt.value == ""){
		alert("미오픈통수를 적어주세요");
		frm.noopencnt.focus();
	}else if(frm.noopenpct.value == ""){
		alert("미오픈율을 적어주세요");
		frm.noopenpct.focus();
	}else{
		frm.submit();
	}
}
</script>

<!-- 표 중간바 시작-->
<table width="100%" cellpadding="1" cellspacing="1" class="a">	
<tr valign="bottom">       
    <td align="left">
		<font color="red"><strong>※ THUNDERMAIL 발송등록</strong></font>
		<br>※ (OFFLINE메일 발송등록시에는 캠페인제목을 OFFLINE으로 고칠 것~!!)<br>
		방법 : textarea의 두번째줄의 텐바이텐관리자 앞을 OFFLINE으로 수정~! 띄어쓰기 해야함(띄어쓰기를 split하기 때문)<br>
		ex)OFFLINE 텐바이텐관리자
    </td>
    <td align="right">
    </td>        
</tr>
</table>
<!-- 표 중간바 끝-->

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form action="" name="frm1" method="post">
<tr bgcolor=#FFFFFF>
	<td>
		<textarea name="ret_txt" cols="120" rows="10"></textarea>
	</td>
	<td>
		<input type="button" value="추출" onclick="javascript:frm1.submit();" class="button">
	</td>
</tr>
</form>
</table>

<% If ret_txt <> "" Then %>
	<table width="100%" border="0" class="a" cellpadding="3" cellspacing="1" bgcolor="#BABABA" align="center">
	<form action="/admin/mailopen/mail_process.asp" name="frm" method="post">
	<input type="hidden" name="mode" value="add">
	<tr bgcolor=#FFFFFF>
		<td align="center">발송구분</td>
		<td align="left">
			<select name="gubun">
				<option>선택하세요</option>
				<option value="mailzine">mailzine</option>
				<option value="fingers">fingers</option>
				<option value="mailzine_not">mailzine_not</option>
				<option value="fingers_not">fingers_not</option>
				<option value="OFFLINE">OFFLINE</option>
			<!-- <option value="academy">academy</option> -->
			</select>
		</td>
	</tr>
	<tr bgcolor=#FFFFFF>
		<td align="center">발송이름</td>
		<td align="left"><input name="title" size="25" type="text" value="<%= ret_txt_seach_re(0) %>_<%=ret_txt_seach_re(11)%>"></td>
	</tr>
	<tr bgcolor=#FFFFFF>
		<td align="center">발송시작시간</td>
		<td align="left"><input name="startdate" size="25" type="text" value="<%= ret_txt_seach_re(11) %>"></td>
	</tr>
	<tr bgcolor=#FFFFFF>
		<td align="center">발송종료시간</td>
		<td align="left"><input name="enddate" size="25" type="text" value="<%= ret_txt_seach_re(21) %>"></td>
	</tr>
	<tr bgcolor=#FFFFFF>
		<td align="center">재발송종료시간</td>
		<td align="left"><input name="reenddate" size="25" type="text" value="<%= ret_txt_seach_re(34) %>"></td>
	</tr>
	<tr bgcolor=#FFFFFF>
		<td align="center">총대상자수</td>
		<td align="left"><input name="totalcnt" size="15" type="text" value="<%= ret_txt_seach_re(4) %>"></td>
	</tr>
	<tr bgcolor=#FFFFFF>
		<td align="center">실발송통수(비율)</td>
		<td align="left"><input name="realcnt" size="15" type="text" value="<%= ret_txt_seach_re(15) %>">
		<input name="realpct" size="10" type="text" value="<%= round((ret_txt_seach_re(15)/ret_txt_seach_re(4))*100,1) %>">
		</td>
	</tr>
	<tr bgcolor=#FFFFFF>
		<td align="center">필터링통수</td>
		<td align="left"><input name="filteringcnt" size="15" type="text" value="<%= ret_txt_seach_re(4)- ret_txt_seach_re(15) %>">
		<input name="filteringpct" size="10" type="text" value="<%= round(((ret_txt_seach_re(4)- ret_txt_seach_re(15))/ret_txt_seach_re(4))*100,1) %>">
		</td>
	</tr>
	<tr bgcolor=#FFFFFF>
		<td align="center">성공발송통수(비율)</td>
		<td align="left"><input name="successcnt" size="15" type="text" value="<%= ret_txt_seach_re(28) %>">
		<input name="successpct" size="10" type="text" value="<%= round((ret_txt_seach_re(28)/ret_txt_seach_re(15))*100,1) %>">
		</td>
	</tr>
	<tr bgcolor=#FFFFFF>
		<td align="center">실패발송통수</td>
		<td align="left"><input name="failcnt" size="15" type="text" value="<%= ret_txt_seach_re(41) %>">
		<input name="failpct" size="10" type="text" value="<%= round((ret_txt_seach_re(41)/ret_txt_seach_re(15))*100,1) %>">
		</td>
	</tr>
	<tr bgcolor=#FFFFFF>
		<td align="center">오픈통수</td>
		<td align="left"><input name="opencnt" size="15" type="text" value="<%= ret_txt_seach_re(56) %>">
		<input name="openpct" size="10" type="text" value="<%= round(( ret_txt_seach_re(56)/ret_txt_seach_re(28))*100,1) %>">
		</td>
	</tr>
	<tr bgcolor=#FFFFFF>
		<td align="center">미오픈통수</td>
		<td align="left"><input name="noopencnt" size="15" type="text" value="<%= ret_txt_seach_re(28)-ret_txt_seach_re(56) %>">
		<input name="noopenpct" size="10" type="text" value="<%= round(((ret_txt_seach_re(28)-ret_txt_seach_re(56))/ret_txt_seach_re(28))*100,1) %>">
		</td>
	</tr>
	<tr bgcolor=#FFFFFF>
		<td align="center">발송메일러</td>
		<td align="left">
			<input type="hidden" name="mailergubun" value="THUNDERMAIL">
			THUNDERMAIL
		</td>
	</tr>
	<!--<tr bgcolor=#FFFFFF>
		<td align="center">추출 내역</td>
		<td align="left">
			<% for i = 0 to 73 %>
				<% response.write ret_txt_seach_re(i)& "_"&i&"<br>" %>
			<% next %>
		</td>
	</tr>-->
	<tr bgcolor=#FFFFFF>
		<td align="center" colspan=2><input type="button" value="저장" onclick="TnMailDataReg();" class="button"></td>
	</tr>
	</form>
	</table>
<% End If %>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
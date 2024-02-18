<%@ Language=VBScript %>
<%
'==========================================================================
'	Description: EMS 서비스지역 조회화면, 서동석
'	History: 2009.04.07
'==========================================================================
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #Include Virtual="/lib/classes/order/clsEms_serviceArea.asp" -->
<%
Dim conListURL	: conListURL = "ems_serviceareaList.asp"
Dim conSaveURL	: conSaveURL = "ems_serviceareaSave.asp"
Dim conViewURL	: conViewURL = "ems_serviceareaView.asp"
Dim conProcURL	: conProcURL = "ems_serviceareaProc.asp"

Dim page		: page			= requestCheckVar(request("page"),10)
Dim CountryCode	: CountryCode	= requestCheckVar(request("CountryCode"),2)
Dim IsUsing		: IsUsing		= requestCheckVar(request("IsUsing"),1)

Dim qString
qString = "CountryCode=" & CountryCode & "&IsUsing=" & IsUsing
conListURL = conListURL & "?" & qString & "&page=" & page
conSaveURL = conSaveURL & "?" & qString & "&page=" & page
conProcURL = conProcURL & "?" & qString & "&page=" & page

rw CountryCode
' 테이블클래스
Dim obj	: Set obj = new clsEms_serviceArea
obj.FRectCountryCode = CountryCode
obj.GetData 

%>

<script language='javascript'>
<!--

// 등록,수정,삭제 처리
function jsSubmit(mode)
{
	var f = document.frmWrite;
	f.mode.value = mode;
	f.submit();
}

//-->
</script>
</head>

<table width="100%" border="0" cellpadding="0" cellspacing="0" style="padding: 5 3 5 10;">
<form name="frmWrite" method="post" action="<%=conProcURL%>">
<input type="hidden" name="mode">
<input type="hidden" name="countryCode" value="<%=obj.FoneItem.FcountryCode%>">		
	
	<tr>
		<td class="td01" align="center">국가명(한글)</td>
		<td class="td02" colspan="3">
			<%=obj.FoneItem.FcountryNameKr%>
		</td>
	</tr>
	<tr>
		<td class="td01" align="center">국가명(영문)</td>
		<td class="td02" colspan="3">
			<%=obj.FoneItem.FcountryNameEn%>
		</td>
	</tr>
	<tr>
		<td class="td01" align="center">EMS요금적용지역</td>
		<td class="td02" colspan="3">
			<%=obj.FoneItem.FemsAreaCode%>
		</td>
	</tr>
	<tr>
		<td class="td01" align="center">EMS최대중량</td>
		<td class="td02" colspan="3">
			<%=obj.FoneItem.FemsMaxWeight%>
		</td>
	</tr>
	<tr>
		<td class="td01" align="center">수취인부담여부</td>
		<td class="td02" colspan="3">
			<%=obj.FoneItem.FreceiverPay%>
		</td>
	</tr>
	<tr>
		<td class="td01" align="center">사용여부</td>
		<td class="td02" colspan="3">
			<%=obj.FoneItem.Fisusing%>
		</td>
	</tr>
	<tr>
		<td class="td01" align="center">기타사항</td>
		<td class="td02" colspan="3">
			<%=obj.FoneItem.FetcContents%>
		</td>
	</tr>

	<tr>
		<td align="center" colspan="4"> 
			<input type="button" class="btnblue" value=" 수 정 " onClick="location.href='<%=conSaveURL%>&countryCode=<%=countryCode%>';">
			&nbsp;&nbsp;
		<%If obj.FoneItem.FisUsing = "Y" Then %>
			<input type="button" class="btnblue" value=" 삭 제 " onClick="jsSubmit('DEL');">
		<%Else %>
			<input type="button" class="btnblue" value=" 사 용 " onClick="jsSubmit('USE');">
		<%End If%>
			&nbsp;&nbsp;
			<input type="button" class="btnblue" value=" 목 록 " onClick="location.href='<%=conListURL%>';">
		</td>
	</tr>
</table>	
</form>
<%
Set obj = Nothing
%>
<!-- 표 하단바 끝-->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->



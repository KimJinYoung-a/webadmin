<%@ Language=VBScript %>
<%
'==========================================================================
'	Description: EMS 서비스지역 등록폼, 서동석
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
Dim conListURL	: conListURL = "ems_weightPrice.asp"
Dim conSaveURL	: conSaveURL = "ems_weightPriceSave.asp"
Dim conProcURL	: conProcURL = "ems_weightPriceProc.asp"

Dim page		: page			    = requestCheckVar(request("page"),10)
Dim scEmsAreaCode	: scEmsAreaCode	= requestCheckVar(request("scEmsAreaCode"),2)
Dim scWeightLimit	: scWeightLimit	= requestCheckVar(request("scWeightLimit"),10)

Dim CompanyCode	: CompanyCode	= requestCheckVar(request("CompanyCode"),3)
Dim emsAreaCode	: emsAreaCode	= requestCheckVar(request("emsAreaCode"),2)
Dim weightLimit	: weightLimit	= requestCheckVar(request("weightLimit"),10)

Dim qString
qString = "scEmsAreaCode=" & scEmsAreaCode & "&scWeightLimit=" & scWeightLimit & "&menupos=" & menupos
conProcURL = conProcURL & "?" & qString & "&page=" & page


' 테이블클래스
Dim obj	: Set obj = new CEms
obj.FRectCompanyCode = CompanyCode
obj.FRectEmsAreaCode = emsAreaCode
obj.FRectWeightLimit = weightLimit
obj.GetWeightPriceData



' 화면표시정보
Dim pageInfo1, pageInfo2, pageInfo3
If countryCode = "" Then
	pageInfo1 = "등록"
	pageInfo2 = "등 록"
Else
	pageInfo1 = "수정"
	pageInfo2 = "수 정"
End If
%>

<script language='javascript'>
<!--

// 등록,수정,삭제 처리
function jsSubmit(mode)
{
	var f = document.frmWrite;
	if (!mode)
		if (f.OEmsAreaCode.value=="")
			f.mode.value = "INS";
		else
			f.mode.value = "UPD";
	else
		f.mode.value = mode;

    if (f.CompanyCode.value.length!=3){
        alert('업체코드 3자를 입력하세요.');
        f.EmsAreaCode.focus();
        return;
    }

    if ((f.EmsAreaCode.value.length!=1) && (f.EmsAreaCode.value.length!=2)) {
        alert('요금적용지역코드 1~2자를 입력하세요.');
        f.EmsAreaCode.focus();
        return;
    }

	if (!validField(f.weightLimit, "중량을"))return ;
	if (!validField(f.emsPrice, "가격을"))return ;

	f.submit();

}

//-->
</script>
</head>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmWrite" method="post" action="<%=conProcURL%>">
<input type="hidden" name="mode">
<input type="hidden" name="OCompanyCode" value="<%= obj.FoneItem.FCompanyCode %>">
<input type="hidden" name="OEmsAreaCode" value="<%= obj.FoneItem.FEmsAreaCode %>">
<input type="hidden" name="OweightLimit" value="<%= obj.FoneItem.FweightLimit %>">
    <tr>
		<td align="center" bgcolor="<%= adminColor("tabletop") %>">업체코드</td>
		<td colspan="3" bgcolor="#FFFFFF">
		<% if obj.FoneItem.FCompanyCode="" then %>
			<input type="text" class="text" name="CompanyCode" value="<%=obj.FoneItem.FCompanyCode%>" size="3" maxlength="3">
		<% else %>
		    <input type="hidden" name="CompanyCode" value="<%=obj.FoneItem.FCompanyCode%>" >
		    <%=obj.FoneItem.FCompanyCode%>
	    <% end if %>
		</td>
	</tr>
	<tr>
		<td align="center" bgcolor="<%= adminColor("tabletop") %>">요금적용지역</td>
		<td colspan="3" bgcolor="#FFFFFF">
		<% if obj.FoneItem.FEmsAreaCode="" then %>
			<input type="text" class="text" name="EmsAreaCode" value="<%=obj.FoneItem.FEmsAreaCode%>" size="2" maxlength="2">
		<% else %>
		    <input type="hidden" name="EmsAreaCode" value="<%=obj.FoneItem.FEmsAreaCode%>" >
		    <%=obj.FoneItem.FEmsAreaCode%>
	    <% end if %>
		</td>
	</tr>
	<tr>
		<td align="center" bgcolor="<%= adminColor("tabletop") %>">중량</td>
		<td colspan="3" bgcolor="#FFFFFF">
		<% if obj.FoneItem.FEmsAreaCode="" then %>
			<input type="text" class="text" name="weightLimit" value="<%=doubleQuote(obj.FoneItem.FweightLimit)%>" size="8" maxlength="10"> (g)까지
		<% else %>
		    <input type="hidden" name="weightLimit" value="<%=obj.FoneItem.FweightLimit%>" >
		    <%=FormatNumber(obj.FoneItem.FweightLimit,0)%> (g)까지
	    <% end if %>


		</td>
	</tr>
	<tr>
		<td align="center" bgcolor="<%= adminColor("tabletop") %>">가격</td>
		<td colspan="3" bgcolor="#FFFFFF">
			<input type="text" class="text" name="emsPrice" value="<%=doubleQuote(obj.FoneItem.FemsPrice)%>" size="8" maxlength="10"> (원)
		</td>
	</tr>


	<tr>
		<td align="center" colspan="4" bgcolor="#FFFFFF">
			<input type="button" class="button" value=" <%=pageInfo2%> " onClick="jsSubmit();">
			&nbsp;&nbsp;
			<% if (obj.FoneItem.FEmsAreaCode<>"") then %>
			<input type="button" class="button" value=" 삭제 " onClick="jsSubmit('DEL');">
			&nbsp;&nbsp;
			<% end if %>
			<input type="button" class="button" value=" 취 소 " onClick="history.back();">
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

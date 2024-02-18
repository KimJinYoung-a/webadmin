<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbSTSopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/maechul/statistic/statisticCls_dw.asp" -->
<!-- #include virtual="/lib/classes/maechul/managementSupport/maechulCls.asp" -->
<!-- #include virtual="/lib/classes/linkedERP/bizSectionCls.asp" -->
<%
	Dim i, cStatistic, vSiteName, vDateGijun, v6MonthDate, vSYear, vSMonth, vSDay, vEYear, vEMonth, vEDay, v6Ago, vbizsec
	dim sellchnl, inc3pl, isSendGift, pggubun
	dim xl

	v6MonthDate	= DateAdd("m",-6,now())
	vSiteName 	= request("sitename")
	vDateGijun	= NullFillWith(request("date_gijun"),"regdate")
	vSYear		= NullFillWith(request("syear"),Year(DateAdd("d",-13,now())))
	vSMonth		= NullFillWith(request("smonth"),Month(DateAdd("d",-13,now())))
	vSDay		= NullFillWith(request("sday"),Day(DateAdd("d",-13,now())))
	vEYear		= NullFillWith(request("eyear"),Year(now))
	vEMonth		= NullFillWith(request("emonth"),Month(now))
	vEDay		= NullFillWith(request("eday"),Day(now))
	v6Ago		= NullFillWith(request("is6ago"),"")
	sellchnl    = requestCheckVar(request("sellchnl"),20)
	vbizsec     = NullFillWith(request("bizsec"),"")
	isSendGift	= requestCheckvar(request("isSendGift"),1)
    inc3pl 		= requestCheckvar(request("inc3pl"),1)
	pggubun		= requestCheckvar(request("pggubun"),2)
	xl 			= requestCheckvar(request("xl"),1)

	Dim vTot_CountPlus, vTot_CountMinus, vTot_MaechulPlus, vTot_MaechulMinus, vTot_Subtotalprice, vTot_Miletotalprice, vTot_MaechulCountSum, vTot_MaechulPriceSum, vTot_sumPaymentEtc

	Set cStatistic = New cStaticTotalClass_list
	cStatistic.FRectDateGijun = vDateGijun
	cStatistic.FRectStartdate = vSYear & "-" & TwoNumber(vSMonth) & "-" & TwoNumber(vSDay)
	cStatistic.FRectEndDate = vEYear & "-" & TwoNumber(vEMonth) & "-" & TwoNumber(vEDay)
	cStatistic.FRectSiteName = vSiteName
	''cStatistic.FRect6MonthAgo = v6Ago
	cStatistic.FRectSellChannelDiv = sellchnl
	''cStatistic.FRectBizSectionCd = vbizsec
	cStatistic.FRectInc3pl = inc3pl  ''2014/01/15 �߰�
	cStatistic.FRectIsSendGift = isSendGift
	cStatistic.FRectPgGubun = pggubun
	cStatistic.fStatistic_dailylist()

    '' �ּ�ó�� 2014/06/23 rdSite ����. Drawsitename �ٽ� �ۼ��ؾ� ��
	''If InStr(vSiteName,"::") > 0 Then
	''	vSiteName = SPlit(vSiteName,"::")(0)
	''End If

if (xl = "Y") then
	Response.Buffer = True
	Response.ContentType = "application/vnd.ms-excel"
	Response.AddHeader "Content-Disposition", "attachment; filename=datamart_on_daily_xl.xls"
else

%>

<script language="javascript">
function image_view(src){
	var image_view = window.open('/admin/culturestation/image_view.asp?image='+src,'image_view','width=1024,height=768,scrollbars=yes,resizable=yes');
	image_view.focus();
}

function searchSubmit()
{
    frm.submit();

    /*
	if((frm.syear.value == <%=Year(v6MonthDate)%> && frm.smonth.value < <%=Month(v6MonthDate)%>) && (frm.is6ago.checked == false))
	{
		alert("6�������� �����ʹ� 6�������������͸� üũ�ϼž� �����մϴ�.");
	}
	else
	{
		frm.submit();
	}
	*/
}

function popXL()
{
    frmXL.submit();
}
</script>

<!-- �˻� ���� -->
<form name="frm" method="get" style="margin:0px;">
<input type="hidden" name="menupos" value="<%= menupos %>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="70" bgcolor="<%= adminColor("gray") %>">�˻� ����</td>
	<td align="left">
		<table class="a">
		<tr>
			<td height="30">
				* �Ⱓ :&nbsp;
				<select name="date_gijun" class="select">
					<option value="regdate" <%=CHKIIF(vDateGijun="regdate","selected","")%>>�ֹ���</option>
					<option value="ipkumdate" <%=CHKIIF(vDateGijun="ipkumdate","selected","")%>>������</option>
				</select>
				<%
					'### ��
					Response.Write "<select name=""syear"" class=""select"">"
					For i=Year(now) To 2001 Step -1
						Response.Write "<option value=""" & i & """ " & CHKIIF(CStr(i)=CStr(vSYear),"selected","") & ">" & i & "</option>"
					Next
					Response.Write "</select>&nbsp;"

					'### ��
					Response.Write "<select name=""smonth"" class=""select"">"
					For i=1 To 12
						Response.Write "<option value=""" & i & """ " & CHKIIF(CStr(i)=CStr(vSMonth),"selected","") & ">" & i & "</option>"
					Next
					Response.Write "</select>&nbsp;"

					'### ��
					Response.Write "<select name=""sday"" class=""select"">"
					For i=1 To 31
						Response.Write "<option value=""" & i & """ " & CHKIIF(CStr(i)=CStr(vSDay),"selected","") & ">" & i & "</option>"
					Next
					Response.Write "</select>&nbsp;~&nbsp;"

					'#############################

					'### ��
					Response.Write "<select name=""eyear"" class=""select"">"
					For i=Year(now) To 2001 Step -1
						Response.Write "<option value=""" & i & """ " & CHKIIF(CStr(i)=CStr(vEYear),"selected","") & ">" & i & "</option>"
					Next
					Response.Write "</select>&nbsp;"

					'### ��
					Response.Write "<select name=""emonth"" class=""select"">"
					For i=1 To 12
						Response.Write "<option value=""" & i & """ " & CHKIIF(CStr(i)=CStr(vEMonth),"selected","") & ">" & i & "</option>"
					Next
					Response.Write "</select>&nbsp;"

					'### ��
					Response.Write "<select name=""eday"" class=""select"">"
					For i=1 To 31
						Response.Write "<option value=""" & i & """ " & CHKIIF(CStr(i)=CStr(vEDay),"selected","") & ">" & i & "</option>"
					Next
					Response.Write "</select>"


					'### 6��������������check
					''Response.Write "<input type=""checkbox"" name=""is6ago"" value=""o"" "
					''If v6Ago = "o" Then
					''	Response.Write "checked"
					''End If
					''Response.Write ">6��������������"
				%>
			</td>
		</tr>
		<tr>
		    <td>
				<%
					'### ����Ʈ����
					Response.Write "* ����Ʈ���� : "
					Call Drawsitename("sitename", vSiteName)

					''Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;* �⺻ ����μ� : "
                    ''Call DrawBizSectionGain("O,T","bizsec", vbizsec,"")
				%>

				    &nbsp;&nbsp;
                	* ä�α���
                	 <% drawSellChannelComboBox "sellchnl",sellchnl %>
				&nbsp;&nbsp;&nbsp;

				<b>* ����ó����</b>
        	    <% Call draw3plMeachulComboBox("inc3pl",inc3pl) %>
				&nbsp;&nbsp;

				<b>* �������</b>
        	    <% Call DrawPggubunName("pggubun",pggubun,"Y","") %>
				&nbsp;&nbsp;

				<label><input type="checkbox" name="isSendGift" value="Y" <%=CHKIIF(isSendGift<>"","checked","")%>>�����ֹ��� ����</label>
			</td>
		</tr>
	    </table>
	</td>
	<td width="110" bgcolor="<%= adminColor("gray") %>"><input type="button" class="button_s" value="�˻�" onClick="javascript:searchSubmit();"></td>
</tr>
</table>
</form>
<!-- �˻� �� -->
<p />

<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="#EEEEEE">
	<tr>
		<td align="left">
			* �˻� �Ⱓ�� ������� ����� �������ϴ�. �׷��� �˻� ��ư�� Ŭ���� �� �ƹ� ������ ����δٰ� ���� �˻���ư�� Ŭ������ ������.
		</td>
		<td align="right">
			<input type="button" class="button" value="�����ޱ�" onClick="popXL()">
		</td>
	</tr>
</table>

<p />

<% end if %>
<table width="100%" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="<%= adminColor("tabletop") %>">
	<td align="center" rowspan="2" colspan="2">�Ⱓ</td>
    <td align="center" colspan="2">�����(+)</td>
    <td align="center" colspan="2">�����(-)</td>
    <td align="center" colspan="2">������հ�</td>
    <td align="center" width="150" rowspan="2">��������<br>(���ϸ��� ����)</td>
	<td align="center" width="150" rowspan="2">���ϸ���</td>
    <td align="center" width="150" rowspan="2">�ǰ����Ѿ�</td>
</tr>
<tr bgcolor="<%= adminColor("tabletop") %>">
    <td align="center">�ֹ��Ǽ�</td>
    <td align="center">�ݾ�</td>
    <td align="center">�ֹ��Ǽ�</td>
    <td align="center">�ݾ�</td>
    <td align="center">�ֹ��Ǽ�</td>
    <td align="center">�ݾ�</td>
</tr>
<%
	For i = 0 To cStatistic.FTotalCount -1
%>
	<tr bgcolor="#FFFFFF">
		<td align="center">
			<% if right(FormatDateTime(cStatistic.flist(i).FRegdate,1),3) = "�����" then %>
				<font color="blue"><%= cStatistic.flist(i).FRegdate %></font>
			<% elseif right(FormatDateTime(cStatistic.flist(i).FRegdate,1),3) = "�Ͽ���" then %>
				<font color="red"><%= cStatistic.flist(i).FRegdate %></font>
			<% else %>
				<%= cStatistic.flist(i).FRegdate %>
			<% end if %>
		</td>
		<td align="center"><%= DateToWeekName(DatePart("w",cStatistic.FList(i).FRegdate)) %></td>
		<td align="center"><%= NullOrCurrFormat(cStatistic.FList(i).FCountPlus) %></td>
		<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).FMaechulPlus) %></td>
		<td align="center"><%= NullOrCurrFormat(cStatistic.FList(i).FCountMinus) %></td>
		<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).FMaechulMinus) %></td>
		<td align="center"><%= NullOrCurrFormat(CLng(cStatistic.FList(i).FCountPlus)+CLng(cStatistic.FList(i).FCountMinus)) %></td>
		<td align="right" style="padding-right:5px;" bgcolor="#E6B9B8"><b><%= NullOrCurrFormat(CLng(cStatistic.FList(i).FMaechulPlus)+CLng(cStatistic.FList(i).FMaechulMinus)) %></b></td>
		<td align="right"><%= NullOrCurrFormat(cStatistic.FList(i).FsumPaymentEtc) %></td>
		<td align="right"><%= NullOrCurrFormat(cStatistic.FList(i).FMiletotalprice) %></td>
		<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).FSubtotalprice - cStatistic.FList(i).FsumPaymentEtc) %></td>
	</tr>
<%
	vTot_CountPlus			= vTot_CountPlus + CLng(NullOrCurrFormat(cStatistic.FList(i).FCountPlus))
	vTot_MaechulPlus		= vTot_MaechulPlus + CLng(NullOrCurrFormat(cStatistic.FList(i).FMaechulPlus))
	vTot_CountMinus			= vTot_CountMinus + CLng(NullOrCurrFormat(cStatistic.FList(i).FCountMinus))
	vTot_MaechulMinus		= vTot_MaechulMinus + CLng(NullOrCurrFormat(cStatistic.FList(i).FMaechulMinus))
	vTot_MaechulCountSum	= vTot_MaechulCountSum + CLng(NullOrCurrFormat(CLng(cStatistic.FList(i).FCountPlus)+CLng(cStatistic.FList(i).FCountMinus)))
	vTot_MaechulPriceSum	= vTot_MaechulPriceSum + CLng(NullOrCurrFormat(CLng(cStatistic.FList(i).FMaechulPlus)+CLng(cStatistic.FList(i).FMaechulMinus)))
	vTot_sumPaymentEtc		= vTot_sumPaymentEtc + CLng(NullOrCurrFormat(cStatistic.FList(i).FsumPaymentEtc))
	vTot_Miletotalprice		= vTot_Miletotalprice + CLng(NullOrCurrFormat(cStatistic.FList(i).FMiletotalprice))
	vTot_Subtotalprice		= vTot_Subtotalprice + CLng(NullOrCurrFormat(cStatistic.FList(i).FSubtotalprice))

	Next
%>
<tr bgcolor="<%= adminColor("tabletop") %>">
	<td align="center" colspan="2">�հ�</td>
	<td align="center"><%=NullOrCurrFormat(vTot_CountPlus)%></td>
	<td align="right" style="padding-right:5px;"><%=NullOrCurrFormat(vTot_MaechulPlus)%></td>
	<td align="center"><%=NullOrCurrFormat(vTot_CountMinus)%></td>
	<td align="right" style="padding-right:5px;"><%=NullOrCurrFormat(vTot_MaechulMinus)%></td>
	<td align="center"><%=NullOrCurrFormat(vTot_MaechulCountSum)%></td>
	<td align="right" style="padding-right:5px;"><b><%=NullOrCurrFormat(vTot_MaechulPriceSum)%></b></td>
	<td align="right"><%=NullOrCurrFormat(vTot_sumPaymentEtc)%></td>
	<td align="right"><%=NullOrCurrFormat(vTot_Miletotalprice)%></td>
	<td align="right" style="padding-right:5px;"><%=NullOrCurrFormat(vTot_Subtotalprice - vTot_sumPaymentEtc)%></td>
</tr>
</table>
<% Set cStatistic = Nothing %>

<form name="frmXL" method="get" style="margin:0px;">
	<input type="hidden" name="xl" value="Y">
	<input type="hidden" name="syear" value="<%= vSYear %>">
	<input type="hidden" name="eyear" value="<%= vEYear %>">
	<input type="hidden" name="smonth" value="<%= vSMonth %>">
	<input type="hidden" name="emonth" value="<%= vEMonth %>">
	<input type="hidden" name="sday" value="<%= vSDay %>">
	<input type="hidden" name="eday" value="<%= vEDay %>">
	<input type="hidden" name="sitename" value="<%= vSiteName %>">
	<input type="hidden" name="sellchnl" value="<%= sellchnl %>">
	<input type="hidden" name="inc3pl" value="<%= inc3pl %>">
</form>
<script src="https://ajax.googleapis.com/ajax/libs/jquery/3.4.1/jquery.min.js"></script>
<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/select2/4.0.13/css/select2.min.css" crossorigin="anonymous" referrerpolicy="no-referrer" />
<script src="https://cdnjs.cloudflare.com/ajax/libs/select2/4.0.13/js/select2.min.js" crossorigin="anonymous" referrerpolicy="no-referrer"></script>
<style type="text/css">
	.select2-container .select2-selection--single {height:17px;}
	.select2-container--default .select2-selection--single .select2-selection__rendered {line-height:16px;}
	.select2-container--default .select2-selection--single .select2-selection__arrow {height: 15px;}
</style>
<script>
$(function() {
	$("select[name=sitename]").select2();
});
</script>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbSTSclose.asp" -->

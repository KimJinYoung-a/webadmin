<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/lib/db/db3Helper.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/PagingCls.asp"-->
<!-- #include virtual="/admin/callStatic/libFunction.asp"-->
<%

dim yyyymmdd_from, yyyymmdd_to, extension, calldate, calldate_from, calldate_to, hour_from, hour_to, phoneno, customerphoneno
dim disposition, lastappsql, pagesize, currpage
dim i, buf

yyyymmdd_from	= requestCheckVar(trim(request.Form("yyyymmdd_from")),10)
yyyymmdd_to		= requestCheckVar(trim(request.Form("yyyymmdd_to")),10)
extension 		= requestCheckVar(trim(request.Form("extension")),3)
hour_from	 	= requestCheckVar(trim(request.Form("hour_from")),2)
hour_to		 	= requestCheckVar(trim(request.Form("hour_to")),2)
disposition		= requestCheckVar(trim(request.Form("disposition")),12)
phoneno			= requestCheckVar(trim(request.Form("phoneno")),12)
customerphoneno	= requestCheckVar(trim(request.Form("customerphoneno")),12)

currpage		= requestCheckVar(trim(request.Form("currpage")),8)
pagesize		= 100



if (yyyymmdd_from = "") then
	yyyymmdd_from = Left((Date - 1), 10)					'�ӽ� - �׽�Ʈ
	yyyymmdd_to = Left((Date - 1), 10)
end if

if (yyyymmdd_from = yyyymmdd_to) then
	if (hour_from <> "") then
		calldate_from 	= yyyymmdd_from & " " & hour_from & ":00:00"
	end if
	if (hour_to <> "") then
		calldate_to 	= yyyymmdd_from & " " & hour_to & ":00:00"
	end if
else
	calldate_from = ""
	calldate_to = ""
end if




if (currpage = "") then
	currpage = 1
end if



Dim strSql
Dim rs

Dim paramInfo
paramInfo = Array(Array("@RETURN_VALUE",adInteger,adParamReturnValue,,0) _
	,Array("@PageSize"		, adInteger	, adParamInput	,		, 50)	_
	,Array("@CurrPage"		, adInteger	, adParamInput	,		, currpage) _
	,Array("@yyyymmdd_from"	, adVarchar	, adParamInput	, 10    , yyyymmdd_from) _
	,Array("@yyyymmdd_to"	, adVarchar	, adParamInput	, 10    , yyyymmdd_to) _
	,Array("@extension" 	, adVarchar	, adParamInput	, 3     , extension) _
	,Array("@calldate_from"	, adVarchar	, adParamInput	, 20    , calldate_from) _
	,Array("@calldate_to"	, adVarchar	, adParamInput	, 20    , calldate_to) _
	,Array("@disposition"	, adVarchar	, adParamInput	, 12    , disposition) _
	,Array("@phoneno"   	, adVarchar	, adParamInput	, 12    , phoneno) _
	,Array("@customerphoneno"	, adVarchar	, adParamInput	, 12    , customerphoneno) _
	,Array("@TotalCount"	, adBigInt	, adParamOutput	,		, 0) _
)

strSql = "db_datamart.dbo.sp_Ten_Call_Inbound_Report"

Call db3_fnExecSPReturnRSOutput(strSql, paramInfo)

If Not db3_rsget.EOF Then
	rs = db3_rsget.getRows()
End If
db3_rsget.close




Dim cPaging
Set cPaging = new PagingCls

cPaging.FTotalCount = GetValue(paramInfo, "@TotalCount")
cPaging.FTotalCount = CInt(cPaging.FTotalCount)
cPaging.FPageSize = 50
cPaging.FCurrPage = currpage
cPaging.Calc

'response.write "----------------" & cPaging.FTotalCount

'response.write "----------------" & dcontext & "----------------"

%>

<script language='javascript'>

function jsPopCal(fName,sName)
{
	var fd = eval("document."+fName+"."+sName);

	if(fd.readOnly==false)
	{
		var winCal;
		winCal = window.open('/lib/common_cal.asp?FN='+fName+'&DN='+sName,'pCal','width=250, height=200');
		winCal.focus();
	}
}

function gotoPage(page)
{
	document.frm.currpage.value = page;
	document.frm.submit();
}

</script>

<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="post" action="">
	<input type="hidden" name="currpage" value="<%= currpage %>">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
		<td align="left">
	       	��¥ : <input type="text" size="10" name="yyyymmdd_from" value="<%=yyyymmdd_from%>" onClick="jsPopCal('frm','yyyymmdd_from');" style="cursor:hand;"> - <input type="text" size="10" name="yyyymmdd_to" value="<%=yyyymmdd_to%>" onClick="jsPopCal('frm','yyyymmdd_to');" style="cursor:hand;"> (������ ��ȭ������ �˻����� �ʽ��ϴ�.)<br>
	       	������ȣ : <% DrawInlinePhoneBox extension %><br>
	       	�ð� : <% DrawCallcenterHourBox hour_from, hour_to %> (1�� �˻��ÿ��� �ð��뺰 �˻��� �����մϴ�.)<br>
	       	<!-- �ϴ� �Ⱥ��̰� �ϰ� �������� ����. ���߿� ������ �𸥴�.
	       	�亯 : <% DrawCallcenterAnswerStateBox disposition %><br>
            -->
            <!--
            ��ȭ��ȣ : <% DrawCallcenterPhoneNameBox phoneno %><br>
            -->
            ����ȣ : <input type="text" size="12" name="customerphoneno" value="<%= customerphoneno %>">
			&nbsp;
		</td>

		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="�˻�" onClick="javascript:document.frm.submit();">
		</td>
	</tr>
	</form>
</table>
<!-- �˻� �� -->
<br>
* ���� : ��� ���� ��ȭ�� ��Ʈ ������ �ý��� ��ȭ�߷�<br>
* ���� : ��� ���� ��ȭ�� ��Ʈ û���� �� ��ȭ����<br>
<!-- ����Ʈ ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
 	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">

        <td>no</td>
        <td>��¥</td>
        <td>�ѰǼ�</td>
        <td>����</td>
        <td>����</td>
        <td>����</td>
        <td>�����ȭ�ð�</td>
        <!--
        <td>�亯</td>
        <td>��������</td>
        -->

</tr>
<%
Dim rowCnt
Dim sRs(20)
dim draw_a, draw_b, draw_c, total_success_count, total_fail_count, total_draw_count, total_call_count, total_success_time, average_success_time

	'buf = "<a href='javascript:alert(""�ٹ��ܽð� �ȳ���Ʈ�� ��ȭ����"")'>����(A)</a>"
	'buf = "<a href='javascript:alert(""���� ��ȭ�� ��ȭ����"")'>����(B)</a>"
'buf = "<a href='javascript:alert(""������ȭ ��ȭ�� ��ȭ����"")'>����(C)</a>"
	'buf = "<a href='javascript:alert(""�ȳ���Ʈ û���� ��ȭ����"")'>����(D)</a>"
'buf = "<a href='javascript:alert(""���� ��ȭ�� ��ȭ����"")'>����(E)</a>"
'buf = "<a href='javascript:alert(""������ȭ ��ȭ�� ��ȭ����"")'>����(F)</a>"

	'buf = "<a href='javascript:alert(""������� - ������??"")'>����(A)</a>"
'buf = "<a href='javascript:alert(""������� - ������??"")'>����(B)</a>"
'buf = "<a href='javascript:alert(""������ȭ - �������"")'>����(C)</a>"
	'buf = "<a href='javascript:alert(""������� - ������??"")'>����(D)</a>"
'buf = "<a href='javascript:alert(""������� - ������??"")'>����(E)</a>"
'buf = "<a href='javascript:alert(""������� - ������??"")'>����(F)</a>"

	'buf = "<a href='javascript:alert(""�� �Է½����� ��ȭ����"")'>����(A)</a>"
	'buf = "<a href='javascript:alert(""�� �Է½����� ��ȭ����"")'>����(B)</a>"
	'buf = "<a href='javascript:alert(""��� ���� ��ȭ�� ��Ʈ û���� ��ȭ����"")'>����(C)</a>"

'����ȭ�Ѽ� = (�亯�Ϸ�Ǽ� + �亯���аǼ� + ��ȭ����Ǽ�)
'�亯�Ϸ�Ǽ� = (count_success_a + count_success_b + count_success_d)
'�亯���аǼ� = (count_fail_a + count_fail_d)
'��ȭ����Ǽ� = (����(A) + ����(B) + ����(C))
'����(A) = totalhangupcount - (count_success_b + count_fail_a + count_success_c + count_fail_b + count_fail_c + count_success_d + count_fail_d)
'����(B) = totaldialcount - (count_success_e + count_fail_e + count_success_f + count_fail_f)
'����(C) = count_draw_c

'���л��� = ������� - ������?? = �亯���аǼ�
'������� = �� �Է½����� ��ȭ���� = ����(A) + ����(B)
'         = ��� ���� ��ȭ�� ��Ʈ û���� ��ȭ���� = ����(C)





'0. yyyymmdd
'1. totalcalltime
'2. totalcallcount

'3. time_success_a
'4. count_success_a

'5. totalhanguptime
'6. totalhangupcount

'7. time_success_b
'8. count_success_b
'9. count_fail_a

'10. time_success_c
'11. count_success_c
'12. count_fail_b

'13. count_fail_c

'14. time_success_d
'15. count_success_d
'16. count_fail_d

'17. totaldialtime
'18. totaldialcount

'19. time_success_e
'20. count_success_e
'21. count_fail_e

'22. time_success_f
'23. count_success_f
'24. count_fail_f

'25. count_draw_c

'26. count_etc

If IsArray(rs) Then
	rowCnt = UBound(rs,2) + 1
%>

	<%For i=0 To UBound(rs,2)%>
    <tr align="center" bgcolor="#FFFFFF">
	<%

		draw_a = rs(6,i) - (rs(8,i) + rs(9,i) + rs(11,i) + rs(12,i) + rs(13,i) + rs(15,i) + rs(16,i))
		draw_b = rs(18,i) - (rs(20,i) + rs(21,i) + rs(23,i) + rs(24,i))
		draw_c = rs(25,i)

		total_fail_count = rs(9,i) + rs(16,i)
		total_draw_count = draw_c' + draw_b + draw_c

		total_call_count = rs(2,i)

		total_success_time = rs(27,i)

		total_success_count = total_call_count - (total_draw_count + total_fail_count)

		if (total_success_count <> 0) then
			average_success_time = total_success_time / total_success_count
		else
			average_success_time = 0
		end if


		' Row �ջ�
		sRs(1) = sRs(1) + 1
		sRs(2) = sRs(2) + CDbl(rs(8,i))

	%>
		<td><%= sRs(1) %></td>
		<td><%= rs(0,i) %></td>
		<td><%= total_call_count %></td>
		<td><%= total_success_count %></td>
		<td><%= total_fail_count %></td>
		<td><%= total_draw_count %></td>

		<td><%= SectoTime(CInt(average_success_time)) %></td>
		<!--
		<td><%= rs(9,i) %></td>
		<td><% PrintCallcenterLastState rs(7,i) %></td>
		-->
	</tr>
	<%Next%>
<!--
    <tr align="center" bgcolor="#FFFFFF">
    	<td><b>�հ�</b></td>
    	<td></td>
		<td></td>
		<td></td>
		<td></td>
		<td></td>
		<td></td>
		<td><<b><%=FormatNumber(sRs(2),0)%></b></td>
		<td></td>-->
		<!--
		<td></td>
		<td></td>
		-->
    <!--</tr>-->
<%
End If
%>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15" align="center">
	   	<% if cPaging.HasPrevScroll then %>
			<span class="list_link"><a href="javascript:gotoPage(<%= cPaging.StartScrollPage-1 %>)">[pre]</a></span>
		<% else %>
		[pre]
		<% end if %>
		<% for i = 0 + cPaging.StartScrollPage to cPaging.StartScrollPage + cPaging.FScrollCount - 1 %>
			<% if (i > cPaging.FTotalpage) then Exit for %>
			<% if CStr(i) = CStr(cPaging.FCurrPage) then %>
			<span class="page_link"><font color="red"><b><%= i %></b></font></span>
			<% else %>
			<a href="javascript:gotoPage(<%= i %>)" class="list_link"><font color="#000000"><%= i %></font></a>
			<% end if %>
		<% next %>
		<% if cPaging.HasNextScroll then %>
			<span class="list_link"><a href="javascript:gotoPage(<%= i %>)">[next]</a></span>
		<% else %>
		[next]
		<% end if %>
	</td>
</tr>
</table>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->

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
dim dcontext, disposition, lastappsql, pagesize, currpage, mode
dim i, buf

yyyymmdd_from	= requestCheckVar(trim(request.Form("yyyymmdd_from")),10)
yyyymmdd_to		= requestCheckVar(trim(request.Form("yyyymmdd_to")),10)
extension 		= requestCheckVar(trim(request.Form("extension")),3)
hour_from	 	= requestCheckVar(trim(request.Form("hour_from")),2)
hour_to		 	= requestCheckVar(trim(request.Form("hour_to")),2)
dcontext		= requestCheckVar(trim(request.Form("dcontext")),40)
disposition		= requestCheckVar(trim(request.Form("disposition")),12)
phoneno			= requestCheckVar(trim(request.Form("phoneno")),12)
customerphoneno	= requestCheckVar(trim(request.Form("customerphoneno")),12)
mode			= requestCheckVar(trim(request.Form("mode")),32)

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




if (mode = "") then
	mode = "all"
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
	,Array("@dcontext"  	, adVarchar	, adParamInput	, 40    , dcontext) _
	,Array("@disposition"	, adVarchar	, adParamInput	, 12    , disposition) _
	,Array("@phoneno"   	, adVarchar	, adParamInput	, 12    , phoneno) _
	,Array("@customerphoneno"	, adVarchar	, adParamInput	, 12    , customerphoneno) _
	,Array("@TotalCount"	, adBigInt	, adParamOutput	,		, 0) _
	,Array("@mode"   		, adVarchar	, adParamInput	, 32    , mode) _
)

strSql = "db_datamart.dbo.sp_Ten_Call_Search"

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
	/*
	if (document.frm.mode.selectedIndex > 0) {
		if ((document.frm.dcontext.selectedIndex != 1) || (document.frm.phoneno.selectedIndex != 1)) {
			alert("��ȭ��ȣ�� �ݼ�����Ʈ�̰�, ���߽��� ������ȭ�϶���\n\n������ȭ �Ǵ� ������ȭ ����Ʈ�� �� �� �ֽ��ϴ�.");
			return;
		}
	}
	*/

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
	       	���߽� : <% DrawCallcenterInOutStateBox dcontext %><br>
            ��ȭ��ȣ : <% DrawCallcenterPhoneNameBox phoneno %><br>
            ����ȣ : <input type="text" size="12" name="customerphoneno" value="<%= customerphoneno %>"><br>
            ���� : <% DrawCallcenterModeBox mode %>
			&nbsp;
		</td>

		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="�˻�" onClick="javascript:gotoPage(1);">
		</td>
	</tr>
	</form>
</table>
<!-- �˻� �� -->
<br>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">
	       	�ȳ���Ʈ(Playback) : �ٹ��� �ð�<br>
	       	��ȭ����(Hangup) : <br>
	       	��ȭ����(Dial) :
		</td>
		<td align="left">
	       	����Ʈ(BackGround) : ��� ���� ��ȭ��<br>
	       	�������(WaitExten) : ���� �������� �����<br>
	       	������(Busy) : ���� ���� �� ��ȭ �ȹ���
		</td>
	</tr>
</table>
<br>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">
	       	�ݼ�����Ʈ(07075490429)<br>
            �繫����Ʈ(07075490556)
		</td>
		<td align="left">
	       	��ǥ��ȣ1(07075490448)<br>
	       	��ǥ��ȣ2(07075490449)<br>
	       	���Ʒ���(07075490559,0216440560)
		</td>
	</tr>
</table>
<br>
�� <%= cPaging.FTotalCount %> ��
<!-- ����Ʈ ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
 	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">

        <td>no</td>
        <td>��¥</td>
        <td>������ȣ</td>
        <td>���̵�</td>
        <td>���߽�</td>
        <td>�߽�</td>
        <td>����</td>
        <td>��ȭ�ð�</td>
        <td>����</td>
        <!--
        <td>�亯</td>
        <td>��������</td>
        -->

</tr>
<%
Dim rowCnt
Dim sRs(20)

'select top 100 yyyymmdd, extension, tenUserID, calldate, src, dst, dcontext, lastapp, duration, disposition, userfield

If IsArray(rs) Then
	rowCnt = UBound(rs,2) + 1
%>

	<%For i=0 To UBound(rs,2)%>
    <tr align="center" bgcolor="#FFFFFF">
	<%
		' Row �ջ�
		sRs(1) = sRs(1) + 1
		sRs(2) = sRs(2) + CDbl(rs(8,i))

	%>
		<td><%= sRs(1) %></td>
		<td><%= rs(3,i) %></td>
		<td><%= rs(1,i) %></td>
		<td><%= rs(2,i) %></td>
		<td><% PrintCallcenterInOutState rs(6,i) %></td>
		<td><% PrintCallcenterPhoneNumberString rs(4,i) %></td>
		<td><% PrintCallcenterPhoneNumberString rs(5,i) %></td>
		<td><%= SectoTime(rs(8,i)) %></td>
		<td>
<%

if ("inbound"=CStr(rs(6,i))) then
	'==========================================================================
	'������ȭ

	if ("Playback"=CStr(rs(7,i))) then
		'======================================================================
		'�ȳ���Ʈ

		buf = "<a href='javascript:alert(""�ٹ��ܽð� �ȳ���Ʈ�� ��ȭ����"")'>����(A)</a>"

	elseif ("Hangup"=CStr(rs(7,i))) then
		'======================================================================
		'��ȭ����

		if ("07075490429"=CStr(rs(5,i))) then
			'==================================================================
			'�ݼ�����Ʈ

			if (rs(8,i) >= 12) then
				buf = "<a href='javascript:alert(""���� ��ȭ�� ��ȭ����"")'>����(B)</a>"
			else
				buf = "<a href='javascript:alert(""������� - ������??"")'>����(A)</a>"
			end if

		elseif ((""<>CStr(rs(1,i))) and (("07075490448"=CStr(rs(5,i))) or ("07075490556"=CStr(rs(5,i))) or ("07075490557"=CStr(rs(5,i))) or ("07075490558"=CStr(rs(5,i))) or ("07075490449"=CStr(rs(5,i))))) then
			'==================================================================
			'������ȭ

			if (rs(8,i) >= 12) then
				buf = "<a href='javascript:alert(""������ȭ ��ȭ�� ��ȭ����"")'>����(C)</a>"
			else
				buf = "<a href='javascript:alert(""������� - ������??"")'>����(B)</a>"
			end if

		elseif ((""=CStr(rs(1,i))) and (("07075490556"=CStr(rs(5,i))) or ("07075490557"=CStr(rs(5,i))) or ("07075490558"=CStr(rs(5,i))))) then
			'==================================================================
			'������ȭ - �������

			buf = "<a href='javascript:alert(""������ȭ - �������"")'>����(C)</a>"

		elseif ((""=CStr(rs(1,i))) and ("07075490448"=CStr(rs(5,i))) or ("07075490449"=CStr(rs(5,i)))) then
			'==================================================================
			'�ٹ��ð��� ����ȭ

			if (rs(8,i) >= 12) then
				buf = "<a href='javascript:alert(""�ȳ���Ʈ û���� ��ȭ����"")'>����(D)</a>"
			else
				buf = "<a href='javascript:alert(""������� - ������??"")'>����(D)</a>"
			end if

		else
			'==================================================================
			'�Է� ����

			buf = "<a href='javascript:alert(""�Է½����� ��ȭ����"")'>����(A)</a>"

		end if

	elseif ("Dial"=CStr(rs(7,i))) then
		'======================================================================
		'��ȭ����

		if ("07075490429"=CStr(rs(5,i))) then
			'==================================================================
			'�ݼ�����Ʈ

			if (rs(8,i) >= 12) then
				buf = "<a href='javascript:alert(""���� ��ȭ�� ��ȭ����"")'>����(E)</a>"
			else
				buf = "<a href='javascript:alert(""������� - ������??"")'>����(E)</a>"
			end if

		elseif ((""<>CStr(rs(1,i))) and (("07075490448"=CStr(rs(5,i))) or ("07075490556"=CStr(rs(5,i))) or ("07075490449"=CStr(rs(5,i))) or ("801"=CStr(rs(5,i))) or ("802"=CStr(rs(5,i))) or ("803"=CStr(rs(5,i))) or ("804"=CStr(rs(5,i))) or ("805"=CStr(rs(5,i))) or ("806"=CStr(rs(5,i))) or ("807"=CStr(rs(5,i))))) then
			'==================================================================
			'������ȭ

			if (rs(8,i) >= 12) then
				buf = "<a href='javascript:alert(""������ȭ ��ȭ�� ��ȭ����"")'>����(F)</a>"
			else
				buf = "<a href='javascript:alert(""������� - ������??"")'>����(F)</a>"
			end if

		else
			'==================================================================
			'�Է� ����

			buf = "<a href='javascript:alert(""�Է½����� ��ȭ����"")'>����(B)</a>"

		end if

	elseif ("BackGround"=CStr(rs(7,i))) then
		'======================================================================
		'����Ʈ : ��� ���� ��ȭ��

		buf = "<a href='javascript:alert(""��� ���� ��ȭ�� ��Ʈ û���� ��ȭ����"")'>����(C)</a>" & CStr(rs(7,i))

	elseif ("WaitExten"=CStr(rs(7,i))) then
		'��������������

		buf = "<a href='javascript:alert(""���� �������� ����� ��ȭ����"")'>����(G)</a>" & CStr(rs(7,i))

	elseif ("Busy"=CStr(rs(7,i))) then
		'��������õ�(toexten) �Ͽ����� ��ȭ �ȹ���

		buf = CStr(rs(7,i))

	else
		'����
		buf = CStr(rs(7,i))
	end if

elseif ("outbound"=CStr(rs(6,i))) then
	'==========================================================================
	'�߽���ȭ
	buf = CStr(rs(7,i))
elseif ("toexten"=CStr(rs(6,i))) then
	'==========================================================================
	'������ȯ
	buf = CStr(rs(7,i))
else
	'==========================================================================
	'����
	buf = CStr(rs(7,i))
end if



'������ȭ
if ("inbound"=CStr(rs(6,i))) then

	if ("ResetCDR"=CStr(rs(7,i))) and (CStr(rs(5,i)) = "1") then
		buf = "<span title='�ݼ��Ϳ� ��ȭ�õ�'><font color=red>�� �õ�</font></span>"
	end if

	if ("Playback"=CStr(rs(7,i))) and (CStr(rs(5,i)) = "07075490449") then
		buf = "<span title='�ٹ��ܽð� �ȳ���Ʈ�� ��ȭ����'><font color=gray>�ȳ���Ʈ1</font></span>"
	end if

	if ("Hangup"=CStr(rs(7,i))) and (CStr(rs(5,i)) = "07075490449") then
		buf = "<span title='�ٹ��ܽð� �ȳ���Ʈ û�� �� ��ȭ����'><font color=gray>�ȳ���Ʈ2</font></span>"
	end if

	if ("BackGround"=CStr(rs(7,i))) and (CStr(rs(5,i)) = "07075490449") then
		buf = "<span title='�ݼ��Ϳ� ��ȭ�õ��Ͽ����� 1 �� ���Է� �Ǵ� �߸� �Է��� �ߴ�'><font color=gray>�õ��ߴ�1</font></span>"
	end if

	if ("WaitExten"=CStr(rs(7,i))) and (CStr(rs(5,i)) = "07075490449") then
		buf = "<span title='�ݼ��Ϳ� ��ȭ�õ��Ͽ����� 1 �� ���Է� �Ǵ� �߸� �Է��� �ߴ�'><font color=gray>�õ��ߴ�2</font></span>"
	end if

	if ("ResetCDR"=CStr(rs(7,i))) and (CStr(rs(5,i)) = "0") then
		buf = "<span title='�繫�ǰ� ����õ�'><font color=black>��ȭ�õ�</font></span>"
	end if

	if ("BackGround"=CStr(rs(7,i))) and (CStr(rs(5,i)) = "07075490448") then
		buf = "<span title='�繫�ǰ� ����õ� �� ������ȣ ���Է� ���·� ��ȭ�ߴ�'><font color=gray>�õ��ߴ�3</font></span>"
	end if

	if ("WaitExten"=CStr(rs(7,i))) and (CStr(rs(5,i)) = "07075490448") then
		buf = "<span title='�繫�ǰ� ����õ� �� ������ȣ ���Է� �Ǵ� �߸��Է� ���·� ��ȭ�ߴ�'><font color=gray>�õ��ߴ�4</font></span>"
	end if

	if ("BackGround"=CStr(rs(7,i))) and (CStr(rs(5,i)) <> "07075490448") and (CStr(rs(5,i)) <> "07075490449") and (CStr(rs(5,i)) <> "1") then
		buf = "<span title='�ݼ��Ϳ� ��ȭ�õ��Ͽ����� 1 �� ���Է� �Ǵ� �߸� �Է��� �ߴ�'><font color=gray>�õ��ߴ�5</font></span>"
	end if

end if



'��������
if ("hunt_context"=CStr(rs(6,i))) then

	if ("Queue"=CStr(rs(7,i))) and (CStr(rs(5,i)) = "07075490429") then
		buf = "<span title='�ݼ��Ϳ� ��ȭ����'><font color=green>�� ����</font></span>"
	end if

	if ("Playback"=CStr(rs(7,i))) and (CStr(rs(5,i)) = "07075490556") then
		buf = "<span title='�繫�� ������ȣ �Է� �� ��ȭ �ȹ���'><font color=gray>��ȭ����</font></span>"
	end if

	if ("Playback"=CStr(rs(7,i))) and (CStr(rs(5,i)) = "07075490559") then
		buf = "<span title='���Ʒ��� �ݼ��� ��ȯ'><font color=gray>���Ʒ���</font></span>"
	end if

	'if ("Queue"=CStr(rs(7,i))) and (CStr(rs(5,i)) = "07075490556") then
	'	buf = "<span title='������ȯ �� �繫�ǰ� ���Ἲ��'><font color=black>��ȭ����</font></span>"
	'end if

	'if ("Dial"=CStr(rs(7,i))) then
	'	''''''''''''''''buf = "<span title='������ȯ �� �繫�ǰ� ��ȭ����'><font color=gray>��ȭ����</font></span>"
	'end if

end if



'��������(���� ���� ��ȭ��ȣ ����??)
if ("pers_context"=CStr(rs(6,i))) then

	if ("Dial"=CStr(rs(7,i))) then
		if (rs(1,i) = rs(5,i)) then
			buf = "<span title='���� ��ȭ��ȣ ���� �õ� �� ��ȭ����'><font color=gray>��ȭ����</font></span>"
		else
			buf = "<span title='���� ��ȭ��ȣ ���� �õ� �� ��ȭ����(���ܹ���:" & rs(1,i) & " > " & rs(5,i) & ")'><font color=gray>��ȭ����</font></span>"
		end if
	end if

	if ("Hangup"=CStr(rs(7,i))) and (rs(5,i) = "908") and (rs(8,i) = 0) then
		buf = "<span title='���� ��ȭ��ȣ ���� �õ� �Ͽ����� ��ȭ����'><font color=red>��ȭ����</font></span>"
	end if

	if ("Hangup"=CStr(rs(7,i))) and (rs(8,i) > 0) then
		buf = "<span title='���� ��ȭ��ȣ ���� �õ� ��ȭ����'><font color=red>��ȭ����</font></span>"
	end if

end if



'�߽���ȭ
if ("outbound"=CStr(rs(6,i))) then

	if (CStr(rs(4,i)) = "0216446030") then
		buf = "<span title='�ݼ��Ϳ��� �ܺο� ��ȭ �õ�'><font color=black>�� ��ȭ</font></span>"
	end if

	if (CStr(rs(4,i)) = "0216440560") then
		buf = "<span title='���Ʒ��� �ݼ��Ϳ��� �ܺο� ��ȭ �õ�'><font color=gray>���Ʒ���</font></span>"
	end if

	if (CStr(rs(4,i)) = "0216441851") then
		buf = "<span title='�������Ϳ��� �ܺο� ��ȭ �õ�'><font color=gray>���� ��ȭ</font></span>"
	end if



end if

response.write buf

%>
		</td>
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

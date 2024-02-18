<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/cscenterv2/lib/incSessionAdminCS.asp" -->
<!-- #include virtual="/cscenterv2/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbACADEMYopen.asp" -->
<!-- #include virtual="/cscenterv2/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->

<%
	Dim i, vOrderSerial, vQuery, vIpkumDiv
	Dim vWantStudyName, vWantStudyYear, vWantStudyMonth, vWantStudyDay, vWantStudyAmPm, vWantStudyHour, vWantStudyMin, vWantStudyPlace, vWantStudyWho
	vOrderSerial = RequestCheckvar(Request("orderserial"),16)

	If vOrderSerial = "" Then
		rw "<script language='javascript'>alert('�߸��� ����Դϴ�.');window.close();</script>"
		Response.End
	Else
		vQuery = "SELECT *, (select ipkumdiv from [db_academy].[dbo].[tbl_academy_order_master] where orderserial = '" & vOrderSerial & "') as ipkumdiv"
		vQuery = vQuery & " FROM [db_academy].[dbo].[tbl_academy_order_weclass] WHERE orderserial = '" & vOrderSerial & "'"
		rsACADEMYget.open vQuery,dbACADEMYget,1
		If Not rsACADEMYget.Eof Then
			vIpkumDiv = rsACADEMYget("ipkumdiv")
			vWantStudyName	= rsACADEMYget("wantstudyName")
			vWantStudyYear	= rsACADEMYget("wantstudyYear")
			vWantStudyMonth	= rsACADEMYget("wantstudyMonth")
			vWantStudyDay	= rsACADEMYget("wantstudyDay")
			vWantStudyAmPm	= rsACADEMYget("wantstudyAmPm")
			vWantStudyHour	= rsACADEMYget("wantstudyHour")
			vWantStudyMin	= rsACADEMYget("wantstudyMin")
			vWantStudyPlace	= rsACADEMYget("wantstudyPlace")
			vWantStudyWho	= rsACADEMYget("wantstudyWho")
		End If
		rsACADEMYget.close
	End If
%>

<script language="javascript">
function goWantSubmit()
{
	if (document.frm1.wantstudyName.value.length<1){
		alert('�ֹ��� ��ü(��ȣȸ)���� �Է��Ͻñ� �ٶ��ϴ�.');
		document.frm1.wantstudyName.focus();
		return;
	}
	if (document.frm1.wantstudyPlace.value.length<1){
		alert('������� �Է��Ͻñ� �ٶ��ϴ�.');
		document.frm1.wantstudyPlace.focus();
		return;
	}
	if (!(document.frm1.wantstudyWho[0].checked) && !(document.frm1.wantstudyWho[1].checked) && !(document.frm1.wantstudyWho[2].checked) && !(document.frm1.wantstudyWho[3].checked)){
		alert('���Ǵ���� �����Ͻñ� �ٶ��ϴ�.');
		return;
	}
	document.frm1.submit();
}
</script>

<table>
<tr>
	<td style="padding:10px 10px 10px 10px;">
		<form name="frm1" action="want_weclass_edit_proc.asp" method="post" style="margin:0px;">
		<input type="hidden" name="orderserial" value="<%= vOrderSerial %>">
		<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
		<tr>
			<td height="25" colspan="5" bgcolor="<%= adminColor("topbar") %>"><img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>��ü���� ��û ���� ����</b></td>
		</tr>
		<tr bgcolor="#FFFFFF">
			<td height="25" width="110" bgcolor="<%= adminColor("topbar") %>">��ü(��ȣȸ)��</td>
			<td><input type="text" name="wantstudyName" style="width:200px;" maxlength="100" value="<%=vWantStudyName%>" /></td>
		</tr>
		<tr bgcolor="#FFFFFF">
			<td height="25" bgcolor="<%= adminColor("topbar") %>">���� �����</td>
			<td>
				<select name="wantstudyYear">
					<option value="2012" <%=CHKIIF(vWantStudyYear="2012","selected","")%>>2012</option>
					<option value="2013" <%=CHKIIF(vWantStudyYear="2013","selected","")%>>2013</option>
					<option value="2014" <%=CHKIIF(vWantStudyYear="2014","selected","")%>>2014</option>
					<option value="2015" <%=CHKIIF(vWantStudyYear="2015","selected","")%>>2015</option>
					<option value="2016" <%=CHKIIF(vWantStudyYear="2016","selected","")%>>2016</option>
					<option value="2017" <%=CHKIIF(vWantStudyYear="2017","selected","")%>>2017</option>
					<option value="2018" <%=CHKIIF(vWantStudyYear="2018","selected","")%>>2018</option>
					<option value="2019" <%=CHKIIF(vWantStudyYear="2019","selected","")%>>2019</option>
					<option value="2020" <%=CHKIIF(vWantStudyYear="2020","selected","")%>>2020</option>
				</select> ��
				<select name="wantstudyMonth">
					<% For i=1 To 12 %>
					<option value="<%=i%>" <%=CHKIIF(CStr(vWantStudyMonth)=CStr(i),"selected","")%>><%=i%></option>
					<% Next %>
				</select> ��
				<select name="wantstudyDay">
					<% For i=1 To 31 %>
					<option value="<%=i%>" <%=CHKIIF(CStr(vWantStudyDay)=CStr(i),"selected","")%>><%=i%></option>
					<% Next %>
				</select> ��
				<select name="wantstudyAmPm">
					<option value="����" <%=CHKIIF(vWantStudyAmPm="����","selected","")%>>����</option>
					<option value="����" <%=CHKIIF(vWantStudyAmPm="����","selected","")%>>����</option>
				</select>
				<select name="wantstudyHour">
					<% For i=1 To 12 %>
					<option value="<%=i%>" <%=CHKIIF(CStr(vWantStudyHour)=CStr(i),"selected","")%>><%=i%></option>
					<% Next %>
				</select> ��
				<select name="wantstudyMin">
					<% For i=0 To 50 step 10 %>
					<option value="<%=i%>" <%=CHKIIF(CStr(vWantStudyMin)=CStr(i),"selected","")%>><%=i%></option>
					<% Next %>
				</select> ��
			</td>
		</tr>
		<tr bgcolor="#FFFFFF">
			<td height="25" bgcolor="<%= adminColor("topbar") %>">�������</td>
			<td><input type="text" name="wantstudyPlace" style="width:450px;" maxlength="100" value="<%=vWantStudyPlace%>" /></td>
		</tr>
		<tr bgcolor="#FFFFFF">
			<td height="25" bgcolor="<%= adminColor("topbar") %>">���Ǵ��</td>
			<td>
				<input name="wantstudyWho" type="radio" value="1" <%=CHKIIF(vWantStudyWho="1","checked","")%> />���&nbsp;
				<input name="wantstudyWho" type="radio" value="2" <%=CHKIIF(vWantStudyWho="2","checked","")%> />��ȣȸ&nbsp;
				<input name="wantstudyWho" type="radio" value="3" <%=CHKIIF(vWantStudyWho="3","checked","")%> />�л�&nbsp;
				<input name="wantstudyWho" type="radio" value="0" <%=CHKIIF(vWantStudyWho="0","checked","")%> />��Ÿ&nbsp;
			</td>
		</tr>
		<tr bgcolor="#FFFFFF">
			<td colspan="2" align="right" style="padding-top:15px;">
				<% If vIpkumdiv = "2" Then %>
				<input type="checkbox" name="gopay" value="o">�����ܰ����� [��û���� MY FINGERS ���� ����]&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
				<% End If %>
				<input type="button" class="button" value="�����ϱ�" onClick="goWantSubmit();">
			</td>
		</tr>
		</table>
	</td>
</tr>
</table>

<!-- #include virtual="/cscenter/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbACADEMYclose.asp" -->
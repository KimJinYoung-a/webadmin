<%
Dim mdlist
Set mdlist = new board
    mdlist.frectprizeyn = "Y"
	mdlist.fnmain_event_MDlist
%>
<script>
if("<%=mdlist.Fresultcount%>" > 0){
	alert("<%=mdlist.Fresultcount%>���� �̵�� ��÷ �̺�Ʈ�� �ֽ��ϴ�!!")
}
function prize(evt_code){
	 var prize = window.open('/admin/eventmanage/event/pop_event_prize.asp?evt_code='+evt_code,'prize','width=800,height=600,scrollbars=yes,resizable=yes');
	 prize.focus();
}
function myblink() {
    document.all.mytext.style.display=document.all.mytext.style.display==""?"none":"";
}
function evtworkerlist(eCode)
{
	var openWorker = null;
	openWorker = window.open('/admin/eventmanage/event/scmMainEvtPopWorkerList.asp?eCode='+eCode+'&team=11','openWorker','width=570,height=570,scrollbars=yes');
	openWorker.focus();
}
setInterval(myblink,500);
</script>


<table width="100%" style="border:1px solid #BABABA" align="center" cellpadding="5" cellspacing="0" class="a">
<form name="frmEvt" style="margin:0px;">
<tr bgcolor="<%= adminColor("tabletop") %>">
    <td>
        <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
		<tr height="25">
		    <td style="border-bottom:1px solid #BABABA">
		        <img src="/images/icon_star.gif" align="absbottom">&nbsp;<b><font color = "red"><span id=mytext>�̺�Ʈ ��÷�� ����Ʈ</span></font></b>
		    </td>
		</tr>
		<% If mdlist.Fresultcount > 0 Then %>
		<tr height="25">
		    <td>
				<table width="100%" border="0" align="center" cellpadding="1" cellspacing="2" class="a">
				<col width = "15%">
				<col />
				<col width = "25%">
				<col width = "10%">
				<tr align="center">
					<td bgcolor="#DCDCDC">�̺�Ʈ��ȣ</td>
					<td bgcolor="#DCDCDC">����</td>
					<td bgcolor="#DCDCDC">�����</td>
					<td bgcolor="#DCDCDC">��÷�Ⱓ</td>
				</tr>
		<% For i = 0 to mdlist.FResultcount -1%>
				<tr align="center">
					<td bgcolor="#EFEFEF"><%=mdlist.FbrdList(i).Fevt_code %></td>
					<td bgcolor="#EFEFEF"><a href="<%=wwwURL%>/event/eventmain.asp?eventid=<%=mdlist.FbrdList(i).Fevt_code%>" target="_blank"><%=mdlist.FbrdList(i).Fevt_name %></a></td>
					<td bgcolor="#EFEFEF">
						<% sbEVTGetwork "selMId",mdlist.FbrdList(i).FpartMDid, "" %>
					</td>
					<td bgcolor="#EFEFEF"><b><font color = "red"><%=mdlist.FbrdList(i).Fevt_laterdate %>��</font></b></td>
				</tr>
		<% Next %>
				</table>
			</td>
		</tr>
		<% Else %>
		<tr height="35">
		    <td align="center">�ش� �̺�Ʈ�� �����ϴ�.</td>
		</tr>
		<% End If %>
        </table>
    </td>
</tr>
</form>
</table>
<% Set mdlist = Nothing %>
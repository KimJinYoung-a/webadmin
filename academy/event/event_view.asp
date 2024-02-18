<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/academy/lib/classes/Event_cls.asp"-->
<%
	'// ���� ���� //
	dim evtId
	dim page, searchKey, searchString, param
	dim spage

	dim oEvent, oPart, i, lp

	'// �Ķ���� ���� //
	evtId = RequestCheckvar(request("evtId"),10)
	page = RequestCheckvar(request("page"),10)
	searchKey = RequestCheckvar(request("searchKey"),16)
	searchString = RequestCheckvar(request("searchString"),128)
	spage = RequestCheckvar(request("spage"),10)

	if spage="" then spage=1
	if searchKey="" then searchKey="evtTitleLong"

	param = "&page=" & page & "&searchKey=" & searchKey & "&searchString=" & searchString	'������ ����

	'// �̺�Ʈ ���� ����
	set oEvent = new CEvent
	oEvent.FRectevtId = evtId

	oEvent.GetNoitceRead


	'// ������ ��� ����
	set oPart = new CPart
	oPart.FRectevtId = evtId
	oPart.FCurrPage = spage

	oPart.GetPartList
%>
<script language="javascript">
<!--
	// �ۻ���
	function GotoEventDel(){
		if (confirm('���� �Ͻðڽ��ϱ�?')){
			document.frm_trans.action="doEvent.asp";
			document.frm_trans.submit();
		}
	}

	// ������ �̵�
	function goPage(pg)
	{
		var frm = document.frm_trans;

		frm.spage.value= pg;
		frm.action="event_view.asp";
		frm.submit();
	}
//-->
</script>
<!-- ���� ȭ�� ���� -->
<table width="100%" border="0" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<tr align="center" bgcolor="#F0F0FD">
	<td colspan="2">
		<table width="100%" border="0" class="a" cellpadding="0" cellspacing="0">
		<tr>
			<td height="26" align="left"><b>�̺�Ʈ �� ����</b></td>
			<td height="26" align="right"><%=oEvent.FEventList(0).Fregdate%>&nbsp;</td>
		</tr>
		</table>
	</td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF">����</td>
	<td bgcolor="#F8F8FF"><%=db2html(oEvent.FEventList(0).FevtTitle)%></td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF">����</td>
	<td bgcolor="#FFFFFF">
	<%
		Select Case oEvent.FEventList(0).FevtType
			Case "M"	'�̹��� �� ����
				if Not(oEvent.FEventList(0).FcontImage="" or isNull(oEvent.FEventList(0).FcontImage)) then
					Response.Write "<img src='" & imgFingers & "/contents/event/" & oEvent.FEventList(0).FcontImage & "' usemap='#evtMainImg' border='0'>"
					Response.Write db2html(oEvent.FEventList(0).FevtCont)
				end if
			Case "H"	'HTML ���۾� ����
				Response.Write db2html(oEvent.FEventList(0).FevtCont)
		end Select
	%>
	</td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF">�Ⱓ</td>
	<td bgcolor="#FFFFFF"><%= FormatDate(oEvent.FEventList(0).FevtSdate,"0000.00.00") & "~" & FormatDate(oEvent.FEventList(0).FevtEdate,"0000.00.00") %></td>
</tr>
<tr><td height="1" colspan="2" bgcolor="#D0D0D0"></td></tr>
<tr>
	<td colspan="2" height="32" bgcolor="#FAFAFA" align="center">
		<img src="/images/icon_modify.jpg" onClick="self.location='Event_modi.asp?menupos=<%=menupos%>&evtId=<%=evtId & param%>'" style="cursor:pointer" align="absmiddle"> &nbsp;
		<img src="/images/icon_delete.gif" onClick="GotoEventDel()" style="cursor:pointer" align="absmiddle"> &nbsp;
		<img src="/images/icon_list.gif" onClick="self.location='Event_list.asp?menupos=<%=menupos & param %>'" style="cursor:pointer" align="absmiddle">
	</td>
</tr>
<form name="frm_trans" method="POST" action="doEvent.asp">
<input type="hidden" name="evtId" value="<%=evtId%>">
<input type="hidden" name="mode" value="delete">
<input type="hidden" name="menupos" value="<%=menupos%>">
<input type="hidden" name="spage" value="<%=spage%>">
<input type="hidden" name="page" value="<%=page%>">
<input type="hidden" name="searchKey" value="<%=searchKey%>">
<input type="hidden" name="searchString" value="<%=searchString%>">
</form>
</table>
<!-- ���� ȭ�� �� -->
<!-- ������ ��� ���� -->
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
	<tr align="center" bgcolor="#F0F0FD">
		<td colspan="8" align="left">��ü�ۼ� : <%= oPart.FTotalCount %> �� Page : <%= spage %>/<%= oPart.FTotalPage %></td>
	</tr>
	<tr align="center" bgcolor="#DDDDFF">
		<td align="center" width="40">��ȣ</td>
		<td align="center" width="80">������</td>
		<td align="center">����</td>
		<td align="center" width="60">������</td>
		<td align="center" width="80">������</td>
		<td align="center" width="60">���ž�(6M)</td>
		<td align="center" width="60">ȸ��������</td>
		<td align="center" width="60">��÷Ƚ��</td>
	</tr>
	<%
		for lp=0 to oPart.FResultCount - 1
	%>
	<tr align="center" bgcolor="#FFFFFF">
		<td><%= oPart.FPartList(lp).FprtId %></td>
		<td><%= oPart.FPartList(lp).FprtUserId & "<br>(" & oPart.FPartList(lp).FprtUserLevel & ")"%></td>
		<td align="left"><%= Replace(db2html(oPart.FPartList(lp).FprtCont2),vbCrLf,"<br>") %></td>
		<td><%= oPart.FPartList(lp).FprtCnt %></td>
		<td><%= FormatDate(oPart.FPartList(lp).FprtDate,"0000.00.00") %></td>
		<td><%= FormatNumber(oPart.FPartList(lp).FsixMonthOrder,0) %></td>
		<td><%= FormatDate(oPart.FPartList(lp).FregDate,"0000.00.00") %></td>
		<td><%= FormatNumber(oPart.FPartList(lp).FprizeCnt,0) %></td>
	</tr>
	<%
		next
	%>
	<tr bgcolor="#FFFFFF">
		<td colspan="8" height="30" align="center">
			<table width="100%" border="0" cellpadding="0" cellspacing="0">
			<tr>
				<td align="center" class="a">
				<!-- ������ ���� -->
				<%
					if oPart.HasPreScroll then
						Response.Write "<a href='javascript:goPage(" & oPart.StarScrollPage-1 & ")'>[pre]</a> &nbsp;"
					else
						Response.Write "[pre] &nbsp;"
					end if
		
					for i=0 + oPart.StarScrollPage to oPart.FScrollCount + oPart.StarScrollPage - 1
		
						if i>oPart.FTotalpage then Exit for
		
						if CStr(spage)=CStr(i) then
							Response.Write " <font color='red'>[" & i & "]</font> "
						else
							Response.Write " <a href='javascript:goPage(" & i & ")'>[" & i & "]</a> "
						end if
		
					next
		
					if oPart.HasNextScroll then
						Response.Write "&nbsp; <a href='javascript:goPage(" & i & ")'>[next]</a>"
					else
						Response.Write "&nbsp; [next]"
					end if
				%>
				<!-- ������ �� -->
				</td>
				<td width="80" align="right">
					<a href="Event_printExcel.asp?evtId=<%=evtId%>"><img src="/images/btn_excel.gif" border="0" align="absmiddle"></a>
				</td>
			</tr>
			</table>
		</td>
	</tr>
</table>
<!-- ������ ��� �� -->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->
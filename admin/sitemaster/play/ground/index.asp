<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/play/playCls.asp" -->
<%
	Dim oground, i , page , state ,idx , title , reservationdate , viewtitle , tagyn , partmdid , partwdid , viewno
	Dim playcate : playcate = 1 'Style+
	menupos = request("menupos")
	page = request("page")
	state = request("state")
	title = request("title")
	reservationdate = request("reservationdate")
	viewtitle = request("viewtitle")
	tagyn = request("tagyn")
	partmdid = request("partmdid")
	partwdid = request("partwdid")
	viewno = request("viewno")

	if page = "" then page = 1

'//�̺�Ʈ ����Ʈ
set oground = new CPlayContents
	oground.FPageSize = 50
	oground.FCurrPage = page
	oground.FRectstate = state
	oground.FRecttitle = viewtitle
	oground.FRPlaycate = playcate
	oground.FRectTag = tagyn
	oground.FRectNo = viewno
	oground.fnGetGroundMainList()
%>
<script type="text/javascript">
	function NextPage(page){
		frm.page.value = page;
		frm.submit();
	}

	function AddNewContents(idx){
		location.href="/admin/sitemaster/play/ground/groundEdit.asp?idx=" + idx;
	}

	function jsSerach(){
		var frm;
		frm = document.frm;
		frm.target = "_self";
		frm.action ="index.asp";
		frm.submit();
	}

	function jsPopCal(sName){
		var winCal;

		winCal = window.open('/lib/common_cal.asp?DN='+sName,'pCal','width=250, height=200');
		winCal.focus();
	}

	function jsTagview(idx){
		var poptag;
		poptag = window.open('/admin/sitemaster/play/lib/pop_tagReg.asp?idx='+idx+'&playcate='+<%=playcate%>,'poptag','width=500,height=400,scrollbars=yes,resizable=yes');
		poptag.focus();
	}
</script>

<form name="frm" method="post" style="margin:0px;">
<input type="hidden" name="page" >
<input type="hidden" name="menupos" value="<%= menupos %>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="<%= adminColor("topbar") %>" >
	<td rowspan="2" width="30" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
	���� : <% Draweventstate2 "state" , state ,"" %>
	&nbsp;&nbsp;&nbsp;
	��ȣ : <input type="text" name="viewno" value="<%=viewno%>" size="5"/>
	<!-- &nbsp;&nbsp;&nbsp;
	������ : <input type="text" name="reservationdate" size=20 maxlength=10 value="<%=reservationdate%>" onClick="jsPopCal('reservationdate');"  style="cursor:pointer;"/> -->
	&nbsp;&nbsp;&nbsp;
	����˻� : <input type="text" name="viewtitle" size=20 value="<%=viewtitle%>" />
	</td>
	<td rowspan="2" width="30" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="javascript:jsSerach();">
	</td>
</tr>
</table>
</form>

<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding:10px 0 10px 0;">
<tr>
	<td align="left">
		<font color="red"> �� ����Ʈ ���� : ���°� ������ �Ͱ� ������ =< ���� �ΰ͸� ������ �˴ϴ�. ������ No. ��ȣ(��������) ������ ����˴ϴ�.</font>
	</td>
	<td align="right">
		<input type="button" class="button" value="�űԵ��" onclick="AddNewContents('0');">
	</td>
</tr>
</table>
<!-- �׼� �� -->

<!-- ����Ʈ ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>" valign="top" border="0">
<tr bgcolor="#FFFFFF">
	<td colspan="20">
		<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
		<tr>
			<td align="left">
				�˻���� : <b><%= oground.FTotalCount%></b>
				&nbsp;
				������ : <b><%= page %> / <%=  oground.FTotalpage %></b>
			</td>
			<td align="right"></td>
		</tr>
		</table>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="5%">idx</td>
	<td width="5%">No.</td>
	<td width="10%">����</td>
	<td width="10%">����(�ڵ�)</td>
	<td width="10%">Ÿ��Ʋ�̹���</td>
	<td width="10%">������</td>
	<td width="10%">�����</td>
	<td width="10%">���WD</td>
	<td width="5%">���</td>
</tr>
<% if oground.FresultCount > 0 then %>
<% for i=0 to oground.FresultCount-1 %>
<tr align="center" bgcolor="#FFFFFF" onmouseout="this.style.backgroundColor='#FFFFFF'" onmouseover="this.style.backgroundColor='#F1F1F1'">
	<td align="center"><%= oground.FItemList(i).Fidx %></td>
	<td align="center"><%= oground.FItemList(i).Fviewno %></td>
	<td align="center"><%= oground.FItemList(i).Fviewtitle %></td>
	<td align="center"><%= geteventstate(oground.FItemList(i).Fstate) %> (<%=oground.FItemList(i).Fstate %>)</td>
	<td align="center"><img src="<%= oground.FItemList(i).Flistimg %>" width=70 border=0></td>
	<td align="center"><%= left(oground.FItemList(i).Freservationdate,10) %></td>
	<td align="center"><%= oground.FItemList(i).FpartMKname %></td>
	<td align="center"><%= oground.FItemList(i).FpartWDname %></td>
	<td align="center">
		<input type="button" class="button" value="����" onclick="AddNewContents('<%= oground.FItemList(i).Fidx %>');"/>
	</td>
</tr>
<% Next %>
<tr>
	<td colspan="20" align="center" bgcolor="#FFFFFF">
	 	<% if oground.HasPreScroll then %>
			<a href="javascript:NextPage('<%= oground.StartScrollPage-1 %>')">[pre]</a>
		<% else %>
			[pre]
		<% end if %>
		<% for i=0 + oground.StartScrollPage to oground.FScrollCount + oground.StartScrollPage - 1 %>
			<% if i>oground.FTotalpage then Exit for %>
			<% if CStr(page)=CStr(i) then %>
			<font color="red">[<%= i %>]</font>
			<% else %>
			<a href="javascript:NextPage('<%= i %>')">[<%= i %>]</a>
			<% end if %>
		<% next %>
		<% if oground.HasNextScroll then %>
			<a href="javascript:NextPage('<%= i %>')">[next]</a>
		<% else %>
			[next]
		<% end if %>
	</td>
</tr>
<% else %>
<tr bgcolor="#FFFFFF">
	<td colspan="20" align="center">[�˻������ �����ϴ�.]</td>
</tr>
<% end if %>
</table>


<%
	set oground = nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->

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
	Dim oPictureDiary, i , page , state ,idx , title , reservationdate , viewtitle , tagyn , viewno
	Dim playcate : playcate = 5
	menupos = request("menupos")
	page = request("page")
	state = request("state")
	title = request("title")
	reservationdate = request("reservationdate")
	viewtitle = request("viewtitle")
	tagyn = request("tagyn")
	viewno = request("viewno")
	
	if page = "" then page = 1

'//�׸��ϱ� ����Ʈ
set oPictureDiary = new CPlayContents
	oPictureDiary.FPageSize = 5
	oPictureDiary.FCurrPage = page
	oPictureDiary.FRectstate = state
	oPictureDiary.FRecttitle = viewtitle
	oPictureDiary.FRPlaycate = playcate
	oPictureDiary.FRectTag = tagyn
	oPictureDiary.FRectNo = viewno
	oPictureDiary.fnGetPictureDiaryList()
%>
<script type="text/javascript">
	function NextPage(page){
		frm.page.value = page;
		frm.submit();
	}

	function AddNewContents(idx){
		var popwin = window.open('/admin/sitemaster/play/picturediary/popPicturediaryEdit.asp?idx=' + idx,'cateHotPosCodeEdit','width=1024,height=768,scrollbars=yes,resizable=yes');
		popwin.focus();
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
		poptag = window.open('/admin/sitemaster/play/lib/pop_tagReg.asp?idx='+idx+'&playcate='+<%=playcate%>,'poptag','width=1024,height=768,scrollbars=yes,resizable=yes');
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
	&nbsp;&nbsp;&nbsp;
	�±� : <select name="tagyn">
				<option value="">��ü</option>
				<option value="Y" <%=chkiif(tagyn="Y","selected","")%>>���</option>
				<option value="N" <%=chkiif(tagyn="N","selected","")%>>�̵��</option>
			 </select>
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
				�˻���� : <b><%= oPictureDiary.FTotalCount%></b>
				&nbsp;
				������ : <b><%= page %> / <%=  oPictureDiary.FTotalpage %></b>
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
	<td width="5%">����(�ڵ�)</td>
	<td width="10%">Ÿ��Ʋ�̹���</td>
	<td width="10%">�±�</td>
	<td width="10%">������</td>
	<td width="3%">���</td>
</tr>
<% if oPictureDiary.FresultCount > 0 then %>
<% for i=0 to oPictureDiary.FresultCount-1 %>
<tr align="center" bgcolor="#FFFFFF" onmouseout="this.style.backgroundColor='#FFFFFF'" onmouseover="this.style.backgroundColor='#F1F1F1'">
	<td align="center"><%= oPictureDiary.FItemList(i).Fidx %></td>
	<td align="center"><%= oPictureDiary.FItemList(i).Fviewno %></td>
	<td align="center"><%= oPictureDiary.FItemList(i).Fviewtitle %></td>
	<td align="center"><%= geteventstate(oPictureDiary.FItemList(i).Fstate) %> (<%=oPictureDiary.FItemList(i).Fstate %>)</td>
	<td align="center"><img src="<%= oPictureDiary.FItemList(i).Flistimg %>" width=70 border=0></td>
	<td align="center"><a href="#" onclick="jsTagview('<%= oPictureDiary.FItemList(i).Fidx %>');" style="cursor:pointer;"><%=chkiif(oPictureDiary.FItemList(i).Ftagcnt>0,"���","�̵��")%>(<%=oPictureDiary.FItemList(i).Ftagcnt%>)</a></td>
	<td align="center"><%= left(oPictureDiary.FItemList(i).Freservationdate,10) %></td>
	<td align="center"><input type="button" class="button" value="����" onclick="AddNewContents('<%= oPictureDiary.FItemList(i).Fidx %>');"/></td>
</tr>
<% Next %>
<tr>
	<td colspan="20" align="center" bgcolor="#FFFFFF">
	 	<% if oPictureDiary.HasPreScroll then %>
			<a href="javascript:NextPage('<%= oPictureDiary.StartScrollPage-1 %>')">[pre]</a>
		<% else %>
			[pre]
		<% end if %>
		<% for i=0 + oPictureDiary.StartScrollPage to oPictureDiary.FScrollCount + oPictureDiary.StartScrollPage - 1 %>
			<% if i>oPictureDiary.FTotalpage then Exit for %>
			<% if CStr(page)=CStr(i) then %>
			<font color="red">[<%= i %>]</font>
			<% else %>
			<a href="javascript:NextPage('<%= i %>')">[<%= i %>]</a>
			<% end if %>
		<% next %>
		<% if oPictureDiary.HasNextScroll then %>
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
	set oPictureDiary = nothing 
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->

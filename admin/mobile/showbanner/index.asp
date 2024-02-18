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
<!-- #include virtual="/admin/mobile/submenu/inc_subhead.asp"-->
<!-- #include virtual="/lib/classes/mobile/showBannerCls.asp" -->
<%
	Dim ostyleplus, i , page , state ,idx , reservationdate , stitle , partmdid , partwdid
	menupos = request("menupos")
	page = request("page")
	state = request("state")
	reservationdate = request("reservationdate")
	stitle = request("stitle")
	partmdid = request("partmdid")
	partwdid = request("partwdid")
	
	if page = "" then page = 1

'//�̺�Ʈ ����Ʈ
set ostyleplus = new CShowBannerContents
	ostyleplus.FPageSize = 50
	ostyleplus.FCurrPage = page
	ostyleplus.FRectstate = state
	ostyleplus.FRecttitle = stitle
	ostyleplus.FRectpartWDid = partwdid
	ostyleplus.FRectpartMDid = partmdid
	ostyleplus.fnGetShowBannerList()
%>
<link href="/js/jqueryui/css/jquery-ui.css" rel="stylesheet">
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/jqueryui/jquery-ui-1.10.2.custom.min.js"></script>
<script type="text/javascript">
	function NextPage(page){
		frm.page.value = page;
		frm.submit();
	}

	function AddNewContents(idx){
		location.href='/admin/mobile/showbanner/popShowbannerEdit.asp?idx=' + idx;
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

	function jsSetItem(idx , styleno){
		var popitem;
		popitem = window.open('/admin/mobile/lib/pop_itemReg.asp?idx='+idx+'&number='+styleno,'popitem','width=600,height=500,scrollbars=yes,resizable=yes');
		popitem.focus();
	}

	$(function(){
		$( "#subList" ).sortable({
			placeholder: "ui-state-highlight",
			start: function(event, ui) {
				ui.placeholder.html('<td height="54" colspan="10" style="border:1px solid #F9BD01;">&nbsp;</td>');
			},
			stop: function(){
				var i=99999;
				$(this).parent().find("input[name^='sort']").each(function(){
					if(i>$(this).val()) i=$(this).val()
				});
				if(i<=0) i=1;
				$(this).parent().find("input[name^='sort']").each(function(){
					$(this).val(i);
					i++;
				});
			}
		});
	});

	function chkAllItem() {
		if($("input[name='chkIdx']:first").attr("checked")=="checked") {
			$("input[name='chkIdx']").attr("checked",false);
		} else {
			$("input[name='chkIdx']").attr("checked","checked");
		}
	}

	function saveList(){
		var chk=0;
		$("form[name='frmlist']").find("input[name='chkIdx']").each(function(){
			if($(this).attr("checked")) chk++;
		});
		if(chk==0) {
			alert("�����Ͻ� �׸��� �������ּ���.");
			return;
		}
		if(confirm("�����Ͻ� ����� ���� ������ �����Ͻðڽ��ϱ�?")) {
			document.frmlist.action="doListModify.asp";
			document.frmlist.submit();
		}
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
	<!-- &nbsp;&nbsp;&nbsp;
	������ : <input type="text" name="reservationdate" size=20 maxlength=10 value="<%=reservationdate%>" onClick="jsPopCal('reservationdate');"  style="cursor:pointer;"/> -->
	&nbsp;&nbsp;&nbsp;
	����˻� : <input type="text" name="stitle" size=20 value="<%=stitle%>" />
	&nbsp;&nbsp;&nbsp;
	�����MD : <% sbGetpartid "partmdid",partmdid,"","11" %>
	&nbsp;&nbsp;&nbsp;
	�����WD : <% sbGetpartid "partwdid",partwdid,"","12" %>
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
<form name="frmlist" method="post" action="">
<input type="hidden" name="menupos" value="<%= request("menupos") %>">
<input type="hidden" name="mode" value="main">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>" valign="top" border="0">
<tr bgcolor="#FFFFFF">
	<td colspan="20">
		<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
		<tr>
			<td align="left">
				�˻���� : <b><%= ostyleplus.FTotalCount%></b>
				&nbsp;
				������ : <b><%= page %> / <%=  ostyleplus.FTotalpage %></b>
				<input type="button" value="��ü����" class="button" onClick="chkAllItem()">
		    	<input type="button" value="��������" class="button" onClick="saveList()" title="ǥ�ü��� �� ��뿩�θ� �ϰ������մϴ�.">
			</td>
			<td align="right"></td>			
		</tr>
		</table>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="5%"></td>
	<td width="5%">idx</td>
	<td width="10%">����</td>
	<td width="5%">���ļ���</td>
	<td width="5%">����(�ڵ�)</td>
	<td width="10%">������</td>
	<td width="10%">�����</td>
	<td width="10%">���WD</td>
	<td width="15%">���</td>
</tr>
<tbody id="subList">
<% if ostyleplus.FresultCount > 0 then %>
<% for i=0 to ostyleplus.FresultCount-1 %>
<tr align="center" bgcolor="#FFFFFF" onmouseout="this.style.backgroundColor='#FFFFFF'" onmouseover="this.style.backgroundColor='#F1F1F1'">
	<td><input type="checkbox" name="chkIdx" value="<%= ostyleplus.FItemList(i).Fidx %>"></td>
	<td align="center"><%= ostyleplus.FItemList(i).Fidx %></td>
	<td align="center"><%= ostyleplus.FItemList(i).Fstitle %></td>
	<td align="center"><input type="text" class="text" name="sort<%=ostyleplus.FItemList(i).Fidx%>" value="<%= ostyleplus.FItemList(i).Fviewno %>" size="2" style="text-align:center"></td>
	<td align="center"><%= geteventstate(ostyleplus.FItemList(i).Fstate) %> (<%=ostyleplus.FItemList(i).Fstate %>)</td>
	<td align="center"><%= left(ostyleplus.FItemList(i).Freservationdate,10) %></td>
	<td align="center"><%= ostyleplus.FItemList(i).FpartMDname %></td>
	<td align="center"><%= ostyleplus.FItemList(i).FpartWDname %></td>
	<td align="center">
		<input type="button" class="button" value="����" onclick="AddNewContents('<%= ostyleplus.FItemList(i).Fidx %>');"/>
		<% If ostyleplus.FItemList(i).Fitemcnt > 0 Then %>
		<input type="button" class="button" value="��ǰȮ��[<%= ostyleplus.FItemList(i).Fitemcnt %>]" onclick="jsSetItem('<%= ostyleplus.FItemList(i).Fidx %>','0');"/>
		<% End If %>
	</td>
</tr>
<% Next %>
</tbody>
<tr>
	<td colspan="20" align="center" bgcolor="#FFFFFF">
	 	<% if ostyleplus.HasPreScroll then %>
			<a href="javascript:NextPage('<%= ostyleplus.StartScrollPage-1 %>')">[pre]</a>
		<% else %>
			[pre]
		<% end if %>
		<% for i=0 + ostyleplus.StartScrollPage to ostyleplus.FScrollCount + ostyleplus.StartScrollPage - 1 %>
			<% if i>ostyleplus.FTotalpage then Exit for %>
			<% if CStr(page)=CStr(i) then %>
			<font color="red">[<%= i %>]</font>
			<% else %>
			<a href="javascript:NextPage('<%= i %>')">[<%= i %>]</a>
			<% end if %>
		<% next %>
		<% if ostyleplus.HasNextScroll then %>
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
</form>


<% 
	set ostyleplus = nothing 
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->

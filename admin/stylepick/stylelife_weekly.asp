<%@ language=vbscript %>
<% option explicit %>
<%
	Response.AddHeader "Cache-Control","no-cache"
	Response.AddHeader "Expires","0"
	Response.AddHeader "Pragma","no-cache"
%>

<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stylepick/stylelifeCls.asp"-->

<%
	Dim oWeekly, cd1,i,page,isusing ,oTheme ,state ,idx , title
	isusing = request("isusing")
	menupos = request("menupos")
	page = request("page")
	state = request("state")
	idx = request("idx")
	title = request("title")
	
	if page = "" then page = 1
	if isusing = "" then isusing = "Y"
		
'//�̺�Ʈ ����Ʈ
set oWeekly = new ClsStyleLife
	oWeekly.FPageSize = 50
	oWeekly.FCurrPage = page
	oWeekly.frectstate = state
	oWeekly.frectidx = idx
	oWeekly.frecttitle = title
	oWeekly.fnGetWeeklyList()
%>

<script language="javascript">

//��ü ����
function jsChkAll(){	
var frm;
frm = document.frm;
	if (frm.chkAll.checked){			      
	   if(typeof(frm.chkitem) !="undefined"){
	   	   if(!frm.chkitem.length){
		   	 	frm.chkitem.checked = true;	   	 
		   }else{
				for(i=0;i<frm.chkitem.length;i++){
					frm.chkitem[i].checked = true;
			 	}		
		   }	
	   }	
	} else {
	  if(typeof(frm.chkitem) !="undefined"){
	  	if(!frm.chkitem.length){
	   	 	frm.chkitem.checked = false;	  
	   	}else{
			for(i=0;i<frm.chkitem.length;i++){
				frm.chkitem[i].checked = false;
			}	
		}		
	  }	
	}
}

// ������ �̵�
function NextPage(ipage){
	document.frm.page.value= ipage;
	document.frm.target = "_self";
	document.frm.action ="stylelife_weekly.asp";
	document.frm.submit();
}

//��� & ����
function reg(idx){
	var weeklyreg = window.open('/admin/stylepick/stylelife_weekly_edit.asp?idx='+idx+'&menupos=<%=menupos%>','weeklyreg','width=1024,height=768,scrollbars=yes,resizable=yes');
	weeklyreg.focus();
}

//����ǰ�߰�
function addnewItem(idx){
	var weeklyitem = window.open('/admin/stylepick/stylelife_weekly_item.asp?idx='+idx+'','weeklyitem','width=500,height=900,scrollbars=yes,resizable=yes');
	weeklyitem.focus();
}

function goReal()
{
	if(confirm("�Ǽ����� �����Ͻðڽ��ϱ�?\n\n�� �����Ͽ� �°� �ֽ� 3�� ��Ŭ���� �������ϴ�.") == true) {
		var stylelifemain = window.open('<%=wwwUrl%>/chtml/stylelife/make_stylelife_main.asp','stylelifemain','width=400,height=300');
		stylelifemain.focus();
	}
}
</script>

<!-- �׼� ���� -->
<form name="frm" method="get" style="margin:0px;">	
<input type="hidden" name="page" >
<input type="hidden" name="idxarr">
<input type="hidden" name="itemcount" value="0">
<input type="hidden" name="mode">
<input type="hidden" name="menupos" value="<%= menupos %>">
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding:10px 0 10px 0;">
<tr>
	<td align="left">
		<input type="button" class="button" value="StyleLife ���� ��� �����ϱ�" onClick="goReal()">
		<font color="red"> �� ����Ʈ ���� : ���°� ������ �Ͱ� ������ =< ���� �ΰ͸� ������ �˴ϴ�. ������ No. ��ȣ(��������) ������ ����˴ϴ�.</font>		
	</td>
	<td align="right">
		<input type="button" class="button" value="�űԵ��" onclick="reg('');">
	</td>
</tr>
</table>
<!-- �׼� �� -->
<br><center><b><font size="5">��Ŭ���۾��� �������� �� ����� �� ��ư(Stylelife���λ�������ϱ�) Ŭ���ϼ���.������ ���� ������ �ȳ��ɴϴ�.</font></b><br><br></center>
<!-- ����Ʈ ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>" valign="top" border="0">
<tr bgcolor="#FFFFFF">
	<td colspan="20">
		<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
		<tr>
			<td align="left">
				�˻���� : <b><%= oWeekly.FTotalCount%></b>
				&nbsp;
				������ : <b><%= page %> /<%=  oWeekly.FTotalpage %></b>
			</td>
			<td align="right">
			</td>			
		</tr>
		</table>
	</td>
	
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td><!--<input type="checkbox" name="chkAll" onClick="jsChkAll();">//--></td>
	<td>No.</td>
	<td>����</td>
	<td>����(�ڵ�)</td>
	<td>Ÿ��Ʋ�̹���</td>
	<td>������</td>
	<td>�����</td>
	<td>��ȹWD</td>
	<td>���</td>
</tr>
<% if oWeekly.FresultCount > 0 then %>
<% for i=0 to oWeekly.FresultCount-1 %>
<tr align="center" bgcolor="#FFFFFF" onmouseout="this.style.backgroundColor='#FFFFFF'" onmouseover="this.style.backgroundColor='#F1F1F1'">
	<td align="center">
		<!--<input type="checkbox" name="chkitem" value="<%= oWeekly.FItemList(i).Fidx %>">//-->
	</td>
	<td align="center">		
		<%= oWeekly.FItemList(i).Fidx %>
	</td>
	<td align="center"><%= oWeekly.FItemList(i).ftitle %></td>
	<td align="center"><%= geteventstate(oWeekly.FItemList(i).fstatename) %> (<%=oWeekly.FItemList(i).fstate %>)</td>
	<td align="center"><img src="<%= oWeekly.FItemList(i).ftitle_img %>" width=200 border=0></td>
	<td align="center"><%= left(oWeekly.FItemList(i).fstartdate,10) %></td>
	<td align="center"><%= oWeekly.FItemList(i).fpartMDname %></td>
	<td align="center"><%= oWeekly.FItemList(i).fpartwDname %></td>
	<td align="center">
		<input type="button" class="button" value="����" onclick="reg('<%= oWeekly.FItemList(i).Fidx %>');">
		<input type="button" value="��ǰ�߰�[<%= oWeekly.FItemList(i).fitemcnt %>]" onclick="addnewItem('<%= oWeekly.FItemList(i).Fidx %>');" class="button">
	</td>
</tr>
<% next %>
<tr>
	<td colspan="20" align="center" bgcolor="#FFFFFF">
	 	<% if oWeekly.HasPreScroll then %>
			<a href="javascript:NextPage('<%= oWeekly.StartScrollPage-1 %>')">[pre]</a>
		<% else %>
			[pre]
		<% end if %>
		<% for i=0 + oWeekly.StartScrollPage to oWeekly.FScrollCount + oWeekly.StartScrollPage - 1 %>
			<% if i>oWeekly.FTotalpage then Exit for %>
			<% if CStr(page)=CStr(i) then %>
			<font color="red">[<%= i %>]</font>
			<% else %>
			<a href="javascript:NextPage('<%= i %>')">[<%= i %>]</a>
			<% end if %>
		<% next %>
		<% if oWeekly.HasNextScroll then %>
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

<% set oWeekly = nothing %>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
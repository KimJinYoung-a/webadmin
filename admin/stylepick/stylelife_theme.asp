<%@ language=vbscript %>
<% option explicit %>
<%
	Response.AddHeader "Cache-Control","no-cache"
	Response.AddHeader "Expires","0"
	Response.AddHeader "Pragma","no-cache"
%>
<%
'###########################################################
' Description : ��Ÿ���� ����
' Hieditor : 2011.04.06 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stylepick/stylelifeCls.asp"-->

<%
Dim cd1,i,page,isusing ,oTheme ,state ,idx , title
	cd1 = request("cd1")
	isusing = request("isusing")
	menupos = request("menupos")
	page = request("page")
	state = request("state")
	idx = request("idx")
	title = request("title")
	
	if page = "" then page = 1
	if isusing = "" then isusing = "Y"
		
'//�̺�Ʈ ����Ʈ
set oTheme = new ClsStyleLife
	oTheme.FPageSize = 50
	oTheme.FCurrPage = page
	oTheme.frectcd1 = cd1
	oTheme.frectstate = state
	oTheme.frectisusing = isusing
	oTheme.frectidx = idx
	oTheme.frecttitle = title
	oTheme.fnGetThemeList()
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

function jsSerach(){
	var frm;
	frm = document.frm;
	frm.target = "_self";
	frm.action ="stylelife_theme.asp";
	frm.submit();
}

// ������ �̵�
function NextPage(ipage){
	document.frm.page.value= ipage;
	document.frm.target = "_self";
	document.frm.action ="stylelife_theme.asp";
	document.frm.submit();
}

//�̺�Ʈ ��� & ����
function eventedit(idx){
	var eventedit = window.open('/admin/stylepick/stylelife_theme_edit.asp?idx='+idx+'&menupos=<%=menupos%>','eventedit','width=1024,height=768,scrollbars=yes,resizable=yes');
	eventedit.focus();
}

//����ǰ�߰�
function addnewItem(idx){
	location.href="/admin/stylepick/stylelife_theme_item.asp?idx="+idx+"&menupos=<%=menupos%>";
}

</script>

<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get">	
<input type="hidden" name="page" >
<input type="hidden" name="idxarr">
<input type="hidden" name="itemcount" value="0">
<input type="hidden" name="mode">
<input type="hidden" name="menupos" value="<%= menupos %>">
<tr align="center" bgcolor="<%= adminColor("topbar") %>" >
	<td rowspan="2" width="30" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
		<select name="cd1" onchange='jsSerach();'>
			<option value="">-��Ÿ��-</option>
			<option value="010" <%=CHKIIF(cd1="010","selected","")%>>Ŭ����</option>
			<option value="020" <%=CHKIIF(cd1="020","selected","")%>>ťƮ</option>
			<option value="040" <%=CHKIIF(cd1="040","selected","")%>>���</option>
			<option value="050" <%=CHKIIF(cd1="050","selected","")%>>���߷�</option>
			<option value="060" <%=CHKIIF(cd1="060","selected","")%>>������Ż</option>
			<option value="070" <%=CHKIIF(cd1="070","selected","")%>>��</option>
			<option value="080" <%=CHKIIF(cd1="080","selected","")%>>�θ�ƽ</option>
			<option value="090" <%=CHKIIF(cd1="090","selected","")%>>��Ƽ��</option>
			<option value="0P0" <%=CHKIIF(cd1="0P0","selected","")%>>��Ÿ����</option>
		</select>
		���� : <% Draweventstate2 "state" , state ," onchange='jsSerach();'" %>		
	</td>	
	<td rowspan="2" width="30" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="javascript:jsSerach();">
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("topbar") %>" >
	<td align="left">
		�׸���ȣ : <input type="text" name="idx" value="<%= idx %>" size=10>
		���� : <input type="text" name="title" value="<%= title %>" size=30>
	</td>
</tr>    
</table>
<br>
<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding:10px 0 10px 0;">
<tr>
	<td align="left">
		<font color="red"> �� ����Ʈ ���� ����(���� �����϶�) : 1. ����(������ȣ��), 2. �׸���ȣ(������ȣ��), 3. ������(�ֱټ�) ������ ����˴ϴ�</font>
		<br>���µǾ��� ��� <b>���� ���ڰ� ���� ������ �ȵ�!!</b> ������.
	</td>
	<td align="right">
		<input type="button" class="button" value="�׸��űԵ��" onclick="eventedit('');">
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
				�˻���� : <b><%= oTheme.FTotalCount%></b>
				&nbsp;
				������ : <b><%= page %> /<%=  oTheme.FTotalpage %></b>				
			</td>
			<td align="right">
			</td>			
		</tr>
		</table>
	</td>
	
</tr>
		
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td><!--<input type="checkbox" name="chkAll" onClick="jsChkAll();">//--></td>
	<td>�׸���ȣ</td>
	<td>��Ÿ��</td>
	<td>����(�ڵ�)</td>
	<td>����̹���</td>
	<td>����</td>
	<td>������</td>
	<td>���³�¥</td>
	<td>��ȹMD</td>
	<td>��ȹWD</td>
	<td>����</td>
	<td>���</td>
</tr>
<% if oTheme.FresultCount > 0 then %>
<% for i=0 to oTheme.FresultCount-1 %>
<tr align="center" bgcolor="#FFFFFF" onmouseout="this.style.backgroundColor='#FFFFFF'" onmouseover="this.style.backgroundColor='#F1F1F1'">
	<td align="center">
		<!--<input type="checkbox" name="chkitem" value="<%= oTheme.FItemList(i).Fidx %>">//-->
	</td>
	<td align="center">		
		<%= oTheme.FItemList(i).Fidx %><br><a href="<%=wwwUrl%>/stylelife/theme/view.asp?idx=<%= oTheme.FItemList(i).Fidx %>&isadmin=admin" onfocus="this.blur()" target="_blink">[�̸�����]</a>
	</td>
	
	<td align="center"><%= CHKIIF(oTheme.FItemList(i).fcatename="","STYLE PICK",oTheme.FItemList(i).fcatename) %></td>
	<td align="center"><%= geteventstate(oTheme.FItemList(i).fstatename) %> (<%=oTheme.FItemList(i).fstate %>)</td>
	<td align="center"><img src="<%= oTheme.FItemList(i).fbanner_img %>" width=50 height=50 border=0></td>
	<td align="center"><%= oTheme.FItemList(i).ftitle %></td>
	<td align="center"><%= left(oTheme.FItemList(i).fstartdate,10) %></td>
	<td align="center">
		<% 
		if oTheme.FItemList(i).fopendate <> "1900-01-01" then response.write oTheme.FItemList(i).fopendate
		'if oTheme.FItemList(i).fclosedate <> "1900-01-01" then response.write oTheme.FItemList(i).fclosedate
		%>
	</td>	
	<td align="center"><%= oTheme.FItemList(i).fpartMDname %></td>
	<td align="center"><%= oTheme.FItemList(i).fpartwDname %></td>
	<td align="center"><%= oTheme.FItemList(i).fsortno %></td>
	<td align="center">
		<input type="button" class="button" value="����" onclick="eventedit('<%= oTheme.FItemList(i).Fidx %>');">
		<input type="button" value="��ǰ�߰�[<%= oTheme.FItemList(i).fitemcnt %>]" onclick="addnewItem('<%= oTheme.FItemList(i).Fidx %>');" class="button">
	</td>
</tr>
<% next %>
<tr>
	<td colspan="20" align="center" bgcolor="#FFFFFF">
	 	<% if oTheme.HasPreScroll then %>
			<a href="javascript:NextPage('<%= oTheme.StartScrollPage-1 %>')">[pre]</a>
		<% else %>
			[pre]
		<% end if %>
		<% for i=0 + oTheme.StartScrollPage to oTheme.FScrollCount + oTheme.StartScrollPage - 1 %>
			<% if i>oTheme.FTotalpage then Exit for %>
			<% if CStr(page)=CStr(i) then %>
			<font color="red">[<%= i %>]</font>
			<% else %>
			<a href="javascript:NextPage('<%= i %>')">[<%= i %>]</a>
			<% end if %>
		<% next %>
		<% if oTheme.HasNextScroll then %>
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
</form>
</table>

<% set oTheme = nothing %>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
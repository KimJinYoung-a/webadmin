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
' Hieditor : 2011.04.07 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stylepick/stylepick_cls.asp"-->

<%
Dim cd1,i,page,isusing ,oevent ,state ,evtidx , title , num
	cd1 = request("cd1")
	isusing = request("isusing")
	menupos = request("menupos")
	page = request("page")
	state = request("state")
	evtidx = request("evtidx")
	title = request("title")
	num = request("num")
	
	if page = "" then page = 1
	if isusing = "" then isusing = "Y"
		
'//�̺�Ʈ ����Ʈ
set oevent = new cstylepick
	oevent.FPageSize = 50
	oevent.FCurrPage = page
	oevent.frectcd1 = cd1
	oevent.frectstate = state
	oevent.frectisusing = isusing
	oevent.frectevtidx = evtidx
	oevent.frecttitle = title
	oevent.fnGetEventList()
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
	frm.action ="stylepick_main_search_event.asp";
	frm.submit();
}

// ������ �̵�
function NextPage(ipage){
	document.frm.page.value= ipage;
	document.frm.target = "_self";
	document.frm.action ="stylepick_main_search_event.asp";
	document.frm.submit();
}

//����ǰ�߰�
function choiceevt(evtidx){
	opener.eval('document.all.divsub'+<%=num%>).innerHTML = "��ȹ���ڵ� & ��ǰ�ڵ� : <input type='text' name='gubunvalue' value='"+evtidx+"' size=10 maxlength=10>";
	self.close();
}
	
</script>

<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="post">	
<input type="hidden" name="page">
<input type="hidden" name="evtidxarr">
<input type="hidden" name="itemcount" value="0">
<input type="hidden" name="mode">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="num" value="<%= num %>">
<tr align="center" bgcolor="<%= adminColor("topbar") %>" >
	<td rowspan="2" width="30" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
		<!--ī�װ� : <%' Drawcategory "cd1",cd1," onchange='jsSerach();'","CD1" %>-->
		<input type="hidden" name="cd1" value="<%= cd1 %>">
		��� : <% drawSelectBoxUsingYN "isusing", isusing %>
		���� : <% Draweventstate2 "state" , state ," onchange='jsSerach();'" %>		
	</td>	
	<td rowspan="2" width="30" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="javascript:jsSerach();">
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("topbar") %>" >
	<td align="left">
		��ȹ����ȣ : <input type="text" name="evtidx" value="<%= evtidx %>" size=10>
		���� : <input type="text" name="title" value="<%= title %>" size=30>
	</td>
</tr>    
</table>
<br>
<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">		
	</td>
	<td align="right">		
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
				�˻���� : <b><%= oevent.FTotalCount%></b>
				&nbsp;
				������ : <b><%= page %> /<%=  oevent.FTotalpage %></b>				
			</td>
			<td align="right">
			</td>			
		</tr>
		</table>
	</td>	
</tr>		
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td><input type="checkbox" name="chkAll" onClick="jsChkAll();"></td>
	<td>��ȹ����ȣ</td>
	<td>ī�װ�</td>
	<td>����(�ڵ�)</td>
	<td>����̹���</td>
	<td>����</td>
	<td>�Ⱓ</td>
	<td>���³�¥<br>���ᳯ¥</td>
	<td>��ȹMD</td>
	<td>��ȹWD</td>
	<td>���</td>
</tr>
<% if oevent.FresultCount > 0 then %>
<% for i=0 to oevent.FresultCount-1 %>
<% if oevent.FItemList(i).fisusing = "Y" then %>
<tr align="center" bgcolor="#FFFFFF" onmouseover=this.style.background="orange"; onmouseout=this.style.background='ffffff';>
<% else %>    
<tr align="center" bgcolor="#FFFFaa" onmouseover=this.style.background="orange"; onmouseout=this.style.background='FFFFaa';>
<% end if %>
	<td align="center">
		<input type="checkbox" name="chkitem" value="<%= oevent.FItemList(i).Fevtidx %>">
	</td>
	<td align="center">
		<%= oevent.FItemList(i).Fevtidx %>
	</td>	
	<td align="center"><%= oevent.FItemList(i).fcatename %></td>
	<td align="center"><%= geteventstate(oevent.FItemList(i).fstatename) %> (<%=oevent.FItemList(i).fstate %>)</td>
	<td align="center"><img src="<%= oevent.FItemList(i).fbanner_img %>" width=50 height=50 border=0></td>
	<td align="center"><%= oevent.FItemList(i).ftitle %></td>
	<td align="center"><%= left(oevent.FItemList(i).fstartdate,10) %><Br>~ <%= left(oevent.FItemList(i).fenddate,10) %></td>
	<td align="center">
		<% 
		if oevent.FItemList(i).fopendate <> "1900-01-01" then response.write oevent.FItemList(i).fopendate
		if oevent.FItemList(i).fclosedate <> "1900-01-01" then response.write oevent.FItemList(i).fclosedate
		%>
	</td>	
	<td align="center"><%= oevent.FItemList(i).fpartMDname %></td>
	<td align="center"><%= oevent.FItemList(i).fpartwDname %></td>
	<td align="center">
		<input type="button" class="button" value="����" onclick="choiceevt('<%= oevent.FItemList(i).Fevtidx %>');">		
	</td>
</tr>
<% next %>
<tr>
	<td colspan="20" align="center" bgcolor="#FFFFFF">
	 	<% if oevent.HasPreScroll then %>
			<a href="javascript:NextPage('<%= oevent.StartScrollPage-1 %>')">[pre]</a>
		<% else %>
			[pre]
		<% end if %>
		<% for i=0 + oevent.StartScrollPage to oevent.FScrollCount + oevent.StartScrollPage - 1 %>
			<% if i>oevent.FTotalpage then Exit for %>
			<% if CStr(page)=CStr(i) then %>
			<font color="red">[<%= i %>]</font>
			<% else %>
			<a href="javascript:NextPage('<%= i %>')">[<%= i %>]</a>
			<% end if %>
		<% next %>
		<% if oevent.HasNextScroll then %>
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

<% set oevent = nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
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
Dim cd1,i,page,isusing ,omain ,state
	cd1 = request("cd1")
	isusing = request("isusing")
	menupos = request("menupos")
	page = request("page")
	state = request("state")
	
	if page = "" then page = 1
	if isusing = "" then isusing = "Y"
		
'//����Ʈ
set omain = new cstylepick
	omain.FPageSize = 50
	omain.FCurrPage = page
	omain.frectcd1 = cd1
	omain.frectstate = state
	omain.frectisusing = isusing
	omain.fnGetmainList()
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
	frm.action ="stylepick_main.asp";
	frm.submit();
}

// ������ �̵�
function NextPage(ipage){
	document.frm.page.value= ipage;
	document.frm.target = "_self";
	document.frm.action ="stylepick_main.asp";
	document.frm.submit();
}

//��� & ����
function mainedit(mainidx){
	var mainedit = window.open('/admin/stylepick/stylepick_main_edit.asp?mainidx='+mainidx+'&menupos=<%=menupos%>','mainedit','width=1024,height=768,scrollbars=yes,resizable=yes');
	mainedit.focus();
}

</script>

<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="post">	
<input type="hidden" name="page" >
<input type="hidden" name="mainidxarr">
<input type="hidden" name="itemcount" value="0">
<input type="hidden" name="mode">
<input type="hidden" name="menupos" value="<%= menupos %>">
<tr align="center" bgcolor="<%= adminColor("topbar") %>" >
	<td rowspan="2" width="30" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
		ī�װ� : <% Drawcategory "cd1",cd1," onchange='jsSerach();'","CD1" %>
		��� : <% drawSelectBoxUsingYN "isusing", isusing %>
		���� : <% Draweventstate2 "state" , state ," onchange='jsSerach();'" %>
	</td>	
	<td rowspan="2" width="30" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="javascript:jsSerach();">
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("topbar") %>" >
	<td align="left">				  	
	</td>
</tr>    
</table>
<br>
<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
		<font color="red"> �ػ��°� "����"�̰� , ���糯¥�� �����Ϻ��� ũ�� ����Ʈ�� �ֱ� ��ϼ����� ����˴ϴ�</font>
	</td>
	<td align="right">
		<input type="button" class="button" value="�űԵ��" onclick="mainedit('');">
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
				�˻���� : <b><%= omain.FTotalCount%></b>
				&nbsp;
				������ : <b><%= page %> /<%=  omain.FTotalpage %></b>				
			</td>
			<td align="right">
			</td>			
		</tr>
		</table>
	</td>	
</tr>		
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td><input type="checkbox" name="chkAll" onClick="jsChkAll();"></td>
	<td>��ȣ</td>
	<td>��Ÿ��</td>
	<td>����(�ڵ�)</td>
	<td>�����̹���</td>
	<td>�Ⱓ</td>
	<td>���³�¥<br>���ᳯ¥</td>
	<td>��ȹMD</td>
	<td>��ȹWD</td>
	<td>���</td>
</tr>
<% if omain.FresultCount > 0 then %>
<% for i=0 to omain.FresultCount-1 %>
<% if omain.FItemList(i).fisusing = "Y" then %>
<tr align="center" bgcolor="#FFFFFF" onmouseover=this.style.background="orange"; onmouseout=this.style.background='ffffff';>
<% else %>    
<tr align="center" bgcolor="#FFFFaa" onmouseover=this.style.background="orange"; onmouseout=this.style.background='FFFFaa';>
<% end if %>
	<td align="center">
		<input type="checkbox" name="chkitem" value="<%= omain.FItemList(i).Fmainidx %>">
	</td>
	<td align="center">
		<a href="/admin/stylepick/index_testview.asp?mainidx=<%= omain.FItemList(i).Fmainidx %>" onfocus="this.blur()" target="_blink">
		<%= omain.FItemList(i).Fmainidx %> [�̸�����]</a>
	</td>	
	<td align="center"><%= omain.FItemList(i).fcatename %></td>
	<td align="center"><%= geteventstate(omain.FItemList(i).fstatename) %> (<%=omain.FItemList(i).fstate %>)</td>
	<td align="center"><img src="<%= omain.FItemList(i).fmainimage %>" width=50 height=50 border=0></td>	
	<td align="center"><%= left(omain.FItemList(i).fstartdate,10) %><Br>~ <%= left(omain.FItemList(i).fenddate,10) %></td>
	<td align="center">
		<% 
		if omain.FItemList(i).fopendate <> "1900-01-01" then response.write omain.FItemList(i).fopendate
		if omain.FItemList(i).fclosedate <> "1900-01-01" then response.write omain.FItemList(i).fclosedate
		%>
	</td>	
	<td align="center"><%= omain.FItemList(i).fpartMDname %></td>
	<td align="center"><%= omain.FItemList(i).fpartwDname %></td>
	<td align="center">
		<input type="button" class="button" value="���� [�۾���<%= omain.FItemList(i).fmaincontentscnt %>��]" onclick="mainedit('<%= omain.FItemList(i).Fmainidx %>');">		
	</td>
</tr>
<% next %>
<tr>
	<td colspan="20" align="center" bgcolor="#FFFFFF">
	 	<% if omain.HasPreScroll then %>
			<a href="javascript:NextPage('<%= omain.StartScrollPage-1 %>')">[pre]</a>
		<% else %>
			[pre]
		<% end if %>
		<% for i=0 + omain.StartScrollPage to omain.FScrollCount + omain.StartScrollPage - 1 %>
			<% if i>omain.FTotalpage then Exit for %>
			<% if CStr(page)=CStr(i) then %>
			<font color="red">[<%= i %>]</font>
			<% else %>
			<a href="javascript:NextPage('<%= i %>')">[<%= i %>]</a>
			<% end if %>
		<% next %>
		<% if omain.HasNextScroll then %>
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

<% set omain = nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
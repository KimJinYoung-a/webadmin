<%@ language=vbscript %>
<% option explicit %>
<%
	Response.AddHeader "Cache-Control","no-cache"
	Response.AddHeader "Expires","0"
	Response.AddHeader "Pragma","no-cache"
%>
<%
'###########################################################
' Description : �÷�Ʈ���� ����
' Hieditor : 2012.03.29 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/items/itemcls_2008.asp"-->
<!-- #include virtual="/lib/classes/color/colortrend_cls.asp"-->

<%
Dim ctcode,i,page,isusing ,ocolor ,state ,iColorCd , partwdid , partmdid , viewno
	iColorCd = request("iCD")
	ctcode = request("ctcode")
	isusing = request("isusing")
	menupos = request("menupos")
	page = request("page")
	state = request("state")
	partwdid = request("partwdid")
	partmdid = request("partmdid")
	viewno = request("viewno")
	
	if page = "" then page = 1
	if isusing = "" then isusing = "Y"

'//����Ʈ
set ocolor = new ccolortrend_list
	ocolor.FPageSize = 50
	ocolor.FCurrPage = page
	ocolor.frectctcode = ctcode
	ocolor.frectcolorcode = iColorCd
	ocolor.frectstate = state
	ocolor.frectisusing = isusing
	ocolor.frectviewno = viewno
	ocolor.frectpartwdid = partwdid
	ocolor.frectpartmdid = partmdid
	ocolor.getcolortrend()
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

function jsSerach(ipage){
	var frm;
	frm = document.frm;
	
	if(frm.ctcode.value!=''){
		if (!IsDouble(frm.ctcode.value)){
			alert('�÷�Ʈ���� �ڵ�� ���ڸ� �����մϴ�.');
			frm.ctcode.focus();
			return;
		}
	}

	frm.page.value= ipage;
	frm.submit();
}

//��� & ����
function popedit(ctcode){
	var popedit = window.open('/admin/itemmaster/colortrend_edit.asp?ctcode='+ctcode+'&menupos=<%=menupos%>','popedit','width=1024,height=768,scrollbars=yes,resizable=yes');
	popedit.focus();
}

//�����ڵ� ����
function selColorChip(cd) {
	document.frm.iCD.value= cd;
	document.frm.submit();
}

</script>

<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="post">	
<input type="hidden" name="page" value=1>
<input type="hidden" name="iCD" value="<%=iColorCd%>">
<input type="hidden" name="mode">
<input type="hidden" name="menupos" value="<%= menupos %>">
<tr align="center" bgcolor="<%= adminColor("topbar") %>" >
	<td rowspan="2" width="30" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
		�÷�Ʈ�����ڵ� : <input type="text" name="ctcode" value="<%=ctcode%>" size=10/>
		&nbsp;No. : <input type="text" name="viewno" size="5" value="<%=viewno%>"/>
		&nbsp;��� : <% drawSelectBoxUsingYN "isusing", isusing %>
		&nbsp;���� : <% Drawcolortrendstate "state" , state ," onchange='jsSerach("""");'" %>
		&nbsp;����� : <% sbGetpartid "partmdid",partmdid,"","23" %>
		&nbsp;�����WD : <% sbGetpartid "partwdid",partwdid,"","12" %>
	</td>	
	<td rowspan="2" width="30" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="javascript:jsSerach('');">
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("topbar") %>" >
	<td align="left">
		<%=FnSelectColorBar(iColorCd,32)%>
	</td>
</tr>    
</table>
<br>
<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
		�ػ��°� "����" �÷��� �ֱ� ������ ����Ʈ�� ������ �÷��� ���� �˴ϴ�.
		<br>���� �������� �÷��� ������� ��¥�� ���� �÷��� �ֱٳ����� ������ �÷��� ���� �˴ϴ�.
	</td>
	<td align="right">
		<input type="button" class="button" value="�űԵ��" onclick="popedit('');">
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
				�˻���� : <b><%= ocolor.FTotalCount%></b>
				&nbsp;
				������ : <b><%= page %> /<%=  ocolor.FTotalpage %></b>				
			</td>
			<td align="right">
			</td>			
		</tr>
		</table>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td><input type="checkbox" name="chkAll" onClick="jsChkAll();"></td>
	<td>�÷�Ʈ����<br>�ڵ�</td>
	<td>No.</td>
	<td>�÷�Ĩ</td>
	<td>����(�ڵ�)</td>
	<td>����</td>
	<td>������</td>
	<td>�ֱټ���</td>
	<td>�����</td>
	<td>�����WD</td>
	<td>���</td>
</tr>
<% if ocolor.FresultCount > 0 then %>
<% for i=0 to ocolor.FresultCount-1 %>
<% if ocolor.FItemList(i).fisusing = "Y" then %>
<tr align="center" bgcolor="#FFFFFF">
<% else %>    
<tr align="center" bgcolor="#FFFFaa">
<% end if %>
	<td align="center">
		<input type="checkbox" name="chkitem" value="<%= ocolor.FItemList(i).fctcode %>">
	</td>
	<td align="center">
		<%= ocolor.FItemList(i).fctcode %>
		<% if ocolor.FItemList(i).fthisweek = ocolor.FItemList(i).fctcode then %>
			&nbsp;&nbsp;<img src="http://fiximage.10x10.co.kr/web2012/colortrend/ico_week.png" width=40 height=40>
		<% end if %>
	</td>
	<td align="center">
		<%= ocolor.FItemList(i).Fviewno %>
	</td>
	<td align="center">
		<img src="<%=ocolor.FItemList(i).FcolorIcon%>" width="20" height="20" alt="<%=ocolor.FItemList(i).fcolorName%>">
	</td>
	<td align="center">
		<%= getcolortrendstate(ocolor.FItemList(i).fstatename) %>
	</td>
	<td align="center"><%=ocolor.FItemList(i).Fcolortitle%></td>
	<td align="center">
		<%= left(ocolor.FItemList(i).fstartdate,10) %>
	</td>	
	<td align="center">
		<%= ocolor.FItemList(i).flastadminid %>
		<Br><%= ocolor.FItemList(i).flastupdate %>	
	</td>
	<td align="center">
		<%= ocolor.FItemList(i).FpartmdName %>
	</td>
	<td align="center">
		<%= ocolor.FItemList(i).FpartwdName %>
	</td>
	<td align="center">
		<input type="button" class="button" value="����" onclick="popedit('<%= ocolor.FItemList(i).fctcode %>');">		
	</td>
</tr>
<% next %>
<tr>
	<td colspan="20" align="center" bgcolor="#FFFFFF">
	 	<% if ocolor.HasPreScroll then %>
			<a href="javascript:jsSerach('<%= ocolor.StartScrollPage-1 %>')">[pre]</a>
		<% else %>
			[pre]
		<% end if %>
		<% for i=0 + ocolor.StartScrollPage to ocolor.FScrollCount + ocolor.StartScrollPage - 1 %>
			<% if i>ocolor.FTotalpage then Exit for %>
			<% if CStr(page)=CStr(i) then %>
			<font color="red">[<%= i %>]</font>
			<% else %>
			<a href="javascript:jsSerach('<%= i %>')">[<%= i %>]</a>
			<% end if %>
		<% next %>
		<% if ocolor.HasNextScroll then %>
			<a href="javascript:jsSerach('<%= i %>')">[next]</a>
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

<% set ocolor = nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
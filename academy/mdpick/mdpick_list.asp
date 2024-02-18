<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/academy/lib/classes/sitemaster/MdpickCls.asp"-->
<%
'###############################################
' PageName : �ΰŽ� ����� ���� mdpick(��! ��õ)
' Discription : ��!��õ(MD pick) ����Ʈ
' History : 2016.08.02 ���¿�
'###############################################

dim page, sDt, eDt, itemid, i, lp, dispCate, SearchUsing , research

research= requestCheckvar(request("research"),2)
page = requestCheckvar(request("page"),16)
if page = "" then page=1
sDt = requestCheckvar(request("sDt"),10)
eDt = requestCheckvar(request("eDt"),10)
itemid = requestCheckvar(request("itemid"),10)
dispCate = requestCheckvar(request("disp"),16)
SearchUsing = requestCheckvar(request("SearchUsing"),1)

if ((research="") and (SearchUsing="")) then 
    SearchUsing = "Y"
end if

dim oMdpick
'set oJust = New Cmdpick
set oMdpick = New Cmdpick
oMdpick.FCurrPage = page
oMdpick.FPageSize=20
oMdpick.FRectSdt = sDt
oMdpick.FRectEdt = eDt
oMdpick.FRectItemId = itemid
oMdpick.FRectIsusing = SearchUsing
oMdpick.FRectDispCate = dispCate
oMdpick.GetMdpick

%>
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script language='javascript'>
function goPage(pg)
{
	document.refreshFrm.page.value=pg;
	document.refreshFrm.action="mdpick_list.asp";
	document.refreshFrm.submit();
}

//��ü ����
function jsChkAll(){
var frm;
frm = document.frmarr;
	if (frm.chkAll.checked){			      
	   if(typeof(frm.cksel) !="undefined"){
	   	   if(!frm.cksel.length){
	   	   	if(frm.cksel.disabled==false){
		   	 	frm.cksel.checked = true;	  
		   	} 	 
		   }else{
				for(i=0;i<frm.cksel.length;i++){
					 	if(frm.cksel[i].disabled==false){
					frm.cksel[i].checked = true;
				}
			 	}		
		   }	
	   }	
	} else {	  
	  if(typeof(frm.cksel) !="undefined"){
	  	if(!frm.cksel.length){
	   	 	frm.cksel.checked = false;	  
	   	}else{
			for(i=0;i<frm.cksel.length;i++){
				frm.cksel[i].checked = false;
			}	
		}		
	  }	
	
	}
	
} 

function ChangeOrderMakerFrame(){ 
	var frm = document.frmarr;
	var upfrm = document.frmArrupdate; 
	var itemcount = 0;
	if(typeof(frm.cksel) !="undefined"){
	 	if(!frm.cksel.length){
	 		if(!frm.cksel.checked){
	 			alert("������ ��ǰ�� �����ϴ�. ��ǰ�� ������ �ּ���");
	 			return;
	 		}
	 		 upfrm.itemid.value = frm.cksel.value;
	 		 itemcount = 1;
	  }else{
	  	for(i=0;i<frm.cksel.length;i++){
	  		if(frm.cksel[i].checked) {	   	    			
	  			if (upfrm.itemid.value==""){
	  			upfrm.itemid.value =  frm.cksel[i].value;
	  			}else{
	  			upfrm.itemid.value =upfrm.itemid.value+ "|" +frm.cksel[i].value;
	  			} 
	  			 itemcount = itemcount+ 1;
	  		}	 
	  	}
	  } 	
	  	if (upfrm.itemid.value == ""){
	  		alert("������ ��ǰ�� �����ϴ�. ��ǰ�� ������ �ּ���");
	 			return;
	  	} 
	}else{
		alert("������ ��ǰ�� �����ϴ�. ��ǰ�� ������ �ּ���");
		return;
	}  

	var ret = confirm('���� ��ǰ�� �����Ͻðڽ��ϱ�?');
	if (ret){
	 upfrm.submit();
		}  
}

</script>
<!-- ��� �˻��� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="refreshFrm" method="get" action="mdpick_list.asp">
<input type="hidden" name="menupos" value="<%= request("menupos") %>">
<input type="hidden" name="page" value="">
<input type="hidden" name="research" value="on">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="80" bgcolor="<%= adminColor("gray") %>">�˻�����</td>
	<td align="left">
		�Ⱓ 
		<input id="sDt" name="sDt" value="<%=sDt%>" class="text" size="10" maxlength="10" /><img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="sDt_trigger" border="0" style="cursor:pointer" align="absmiddle" /> ~
		<input id="eDt" name="eDt" value="<%=eDt%>" class="text" size="10" maxlength="10" /><img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="eDt_trigger" border="0" style="cursor:pointer" align="absmiddle" /> /
		<script language="javascript">
			var CAL_Start = new Calendar({
				inputField : "sDt", trigger    : "sDt_trigger",
				onSelect: function() {
					var date = Calendar.intToDate(this.selection.get());
					CAL_End.args.min = date;
					CAL_End.redraw();
					this.hide();
				}, bottomBar: true, dateFormat: "%Y-%m-%d"
			});
			var CAL_End = new Calendar({
				inputField : "eDt", trigger    : "eDt_trigger",
				onSelect: function() {
					var date = Calendar.intToDate(this.selection.get());
					CAL_Start.args.max = date;
					CAL_Start.redraw();
					this.hide();
				}, bottomBar: true, dateFormat: "%Y-%m-%d"
			});
		</script>
		��ǰ�ڵ� <input type="text" name="itemid" class="text" size="12" value="<%=itemid%>">
		&nbsp;
		����ī�װ�: <!-- #include virtual="/academy/comm/dispCateSelectBox.asp"--> 

		<b> �� �� : </b>
		<select name="SearchUsing">
			<option value ="" style="color:blue">�� ü</option>
			<option value="Y" <% If "Y" = cstr(SearchUsing) Then%> selected <%End if%>>Y</option>
			<option value="N" <% If "N" = cstr(SearchUsing) Then%> selected <%End if%>>N</option>
		</select>&nbsp;&nbsp;&nbsp;
		
	</td>
	<td width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="submit" class="button_s" value="�˻�">
	</td>
</tr>
</form>
</table>
<!-- �˻� �� -->
<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding:10 0 10 0;">
<form name="frmarr" method="post" action="doMdpick_Process.asp">
<input type="hidden" name="menupos" value="<%= request("menupos") %>">
<input type="hidden" name="mode" value="">
<tr>
	<td align="right"><input type="button" value="������ �߰�" onclick="self.location='Mdpick_write.asp?mode=add&menupos=<%= menupos %>'" class="button"></td>
</tr>
</table>
<!-- �׼� �� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="10" >
		�˻���� : <b><%=oMdpick.FtotalCount%></b>
		&nbsp;
		������ : <b><%= page %> / <%=oMdpick.FtotalPage%></b>
	</td>
</tr>

<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="30">
		<input type="button" class="button" value="����" onClick="ChangeOrderMakerFrame()">
		<input type="checkbox" name="chkAll" onClick="jsChkAll();">
	</td>
	<td>IDX</td>
	<td>������</td>
	<td>������</td>
	<td>Image</td>
	<td>[��ǰ�ڵ�] ��ǰ��</td>
	<td>����ī�װ�</td>
	<td>ǰ��</td>
	<td>��뿩��</td>
	<td colspan="2">�����</td>
</tr>
<%	if oMdpick.FResultCount < 1 then %>
<tr>
	<td colspan="10" height="60" align="center" bgcolor="#FFFFFF">���(�˻�)�� �������� �����ϴ�.</td>
</tr>
<%
	else
		for i=0 to oMdpick.FResultCount-1
%>
<a href="mdpick_write.asp?mode=edit&menupos=<%= menupos %>&idx=<%= oMdpick.FItemList(i).Fidx %>">
<tr <% if cstr(oMdpick.FItemList(i).Fenddate) < cstr(date()) or oMdpick.FItemList(i).Fisusing="N" then %>bgcolor="<%= adminColor("dgray") %>"<% else %>bgcolor="#FFFFFF" style="cursor:pointer;" onmouseover=this.style.background="f1f1f1"; onmouseout=this.style.background='ffffff';<% end if %> >
	<td align="center"  width="30"> 
		<input type="checkbox" name="cksel" value="<%= oMdpick.FItemList(i).Fidx %>" <% If (oMdpick.FItemList(i).FIsusing = "N") then %>disabled<% End if %>>
	</td>
	<td align="center"><%= oMdpick.FItemList(i).Fidx %></td>
	<td align="center"><%= oMdpick.FItemList(i).Fstartdate %></td>
	<td align="center"><%= oMdpick.FItemList(i).Fenddate %></td>
	<td align="center"><a href="mdpick_write.asp?mode=edit&menupos=<%= menupos %>&idx=<%= oMdpick.FItemList(i).Fidx %>"><img src="<%= oMdpick.FItemList(i).FsmallImage %>" width="50" height="50" border="0"></a></td>
	<td align="center"><a href="<%= wwwFingers %>/diyshop/shop_prd.asp?itemid=<%=oMdpick.FItemList(i).FItemID%>" target="_blank"><font color="blue"><%= "[" & oMdpick.FItemList(i).FItemID & "] " %></font></a> <%= oMdpick.FItemList(i).FItemname %></td>
	<td align="center"><%=fnCateCodeNameSplit(oMdpick.FItemList(i).FCateName,oMdpick.FItemList(i).FItemID)%></span></td>
	<td align="center"><% if oMdpick.FItemList(i).FsellYn<>"Y" then Response.Write "ǰ��" %></td>
	<td align="center"><%= oMdpick.FItemList(i).Fisusing %></td>
	<td align="center" colspan="2"><%= left(oMdpick.FItemList(i).Fregdate,10) %></td>
</tr>
</a>
<%
		next
	end if
%>
<!-- ���� ��� �� -->
<tr bgcolor="#FFFFFF">
	<td colspan="10" align="center">
	<!-- ������ ���� -->
	<%
		if oMdpick.HasPreScroll then
			Response.Write "<a href='javascript:goPage(" & oMdpick.StartScrollPage-1 & ")'>[pre]</a> &nbsp;"
		else
			Response.Write "[pre] &nbsp;"
		end if

		for lp=0 + oMdpick.StartScrollPage to oMdpick.FScrollCount + oMdpick.StartScrollPage - 1

			if lp>oMdpick.FTotalpage then Exit for

			if CStr(page)=CStr(lp) then
				Response.Write " <font color='red'>" & lp & "</font> "
			else
				Response.Write " <a href='javascript:goPage(" & lp & ")'>" & lp & "</a> "
			end if

		next

		if oMdpick.HasNextScroll then
			Response.Write "&nbsp; <a href='javascript:goPage(" & lp & ")'>[next]</a>"
		else
			Response.Write "&nbsp; [next]"
		end if
	%>
	<!-- ������ �� -->
	</td>
</tr>
</form>
</table>
<form name="frmArrupdate" method="post" action="delmdpickarr.asp">
<input type="hidden" name="mode" value="del">
<input type="hidden" name="itemid" value="">
</form>
<%
set oMdpick = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->
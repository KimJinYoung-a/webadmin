<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ���̶�� ��ǰ ī�װ� ����
' Hieditor : 2013.05.10 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/ithinkso/category/category_cls_ithinkso.asp"-->

<%
Dim oitem, i, page, itemid, itemname, makerid,CateSeq0, CateSeq1, CateSeq2, CateSeq3, sellyn, usingyn
dim Depth, cCate, reload
	CateSeq0	=	requestCheckVar(Request("iCateSeq0"),10)
	CateSeq1	=	requestCheckVar(Request("iCateSeq1"),10)
	CateSeq2	=	requestCheckVar(Request("iCateSeq2"),10)
	CateSeq3	=	requestCheckVar(Request("iCateSeq3"),10)
	Depth 		= 	requestCheckVar(Request("Depth"),10)
	page = request("page")
	reload = request("reload")
	itemid		= request("itemid")
	itemname	= request("itemname")
	makerid		= request("makerid")
	sellyn		= request("sellyn")
	usingyn		= request("usingyn")

if page = "" then page = 1
IF Depth = "" THEN Depth = 0
if reload = "" and usingyn = "" then usingyn = "Y"

'if CateSeq0 = "" then CateSeq0 = 1
if CateSeq0 = "" then
	CateSeq1 = ""
	CateSeq2 = ""
	CateSeq3 = ""
elseif CateSeq1 = "" then
	CateSeq2 = ""
	CateSeq3 = ""
elseif CateSeq2 = "" then
	CateSeq3 = ""	
end if

set oitem = new ccategory_ithinkso
	oitem.FPageSize         = 50
	oitem.FCurrPage         = page
	oitem.FRectMakerid      = makerid
	oitem.FRectSellYN       = sellyn
	oitem.FRectIsUsing      = usingyn
	oitem.FRectItemid       = itemid
	oitem.FRectItemName     = itemname
	oitem.frectCateTypeSeq   = CateSeq0
	oitem.FRectCateSeq1   = CateSeq1
	oitem.FRectCateSeq2     = CateSeq2
	oitem.FRectCateSeq3   = CateSeq3
	oitem.getCategoryitem
%>
<script type="text/javascript">

	//�űԵ�� �˾�
	function categoryitemreg(){
	
		//-- ī�װ� ------------------------------------------
		if(frmSearch.CateSeq0.options[frmSearch.CateSeq0.selectedIndex].value ==""){
			alert("����Ͻ� ī�װ� Ÿ���� �����ϼ���");
			document.frmSearch.CateSeq0.focus();
			return;
		}
		
		if(frmSearch.CateSeq1.options[frmSearch.CateSeq1.selectedIndex].value ==""){
			alert("����Ͻ� ��ī�װ��� �����ϼ���");
			document.frmSearch.CateSeq1.focus();
			return;
		}

		if(frmSearch.CateSeq2.options[frmSearch.CateSeq2.selectedIndex].value ==""){
			alert("����Ͻ� ��ī�װ��� �����ϼ���");
			document.frmSearch.CateSeq2.focus();
			return;
		}

		if(frmSearch.CateSeq3.options[frmSearch.CateSeq3.selectedIndex].value ==""){
			if (confirm('��ī�װ��� ������ �Ǿ� ���� �ʽ��ϴ�. �����Ͻðڽ��ϱ�?')){			
				document.frmSearch.CateSeq3.focus();
			}else{
				return;
			}	
		}
				
		var CateSeq0 = document.frmSearch.CateSeq0.options[frmSearch.CateSeq0.selectedIndex].value
		var CateSeq1 = document.frmSearch.CateSeq1.options[frmSearch.CateSeq1.selectedIndex].value
		var CateSeq2 = document.frmSearch.CateSeq2.options[frmSearch.CateSeq2.selectedIndex].value
		var CateSeq3 = document.frmSearch.CateSeq3.options[frmSearch.CateSeq3.selectedIndex].value

		var categoryitemreg = window.open('/admin/ithinkso/category/category_item_reg_ithinkso.asp?CateSeq0='+CateSeq0+'&CateSeq1='+CateSeq1+'&CateSeq2='+CateSeq2+'&CateSeq3='+CateSeq3,'categoryitemreg','width=1024,height=768,scrollbars=yes,resizable=yes');
		categoryitemreg.focus();
	}

	function frmsubmit(page){
		frmSearch.page.value=page;
		frmSearch.submit();
	}
	
	//���� ī�װ� ���濡 ���� ���� ī�װ� ������ ���� ó��
	function jsChCategory(intD){				
		var intT = 0;	
		eval("document.frmSearch.iCateSeq"+intD).value =  eval("document.frmSearch.CateSeq"+intD).options[eval("document.frmSearch.CateSeq"+intD).selectedIndex].value;					
		if(eval("document.frmSearch.iCateSeq"+intD).value ==""){
		  if (intD == 0) {
		    document.frmSearch.Depth.value="";
		    frmsubmit('');	
		  }else{
			jsChCategory(intD-1);
		  }
		}else{
			intT= eval("document.frmSearch.CateSeq"+intD).options[eval("document.frmSearch.CateSeq"+intD).selectedIndex].thread;		
									
			document.frmSearch.Depth.value = intD;
			
			frmsubmit('');		
		}	
	}

	//ī�װ���ǰ����
	function categoryitemdel(upfrm){

		if (!CheckSelected()){
				alert('���þ������� �����ϴ�.');
				return;
			}	
			var frm;
				for (var i=0;i<document.forms.length;i++){
					frm = document.forms[i];
					if (frm.name.substr(0,9)=="frmBuyPrc") {
						if (frm.cksel.checked){
							upfrm.CateDispSeqarr.value = upfrm.CateDispSeqarr.value + frm.CateDispSeq.value + ','
								
						}
					}
				}
				
		upfrm.target = "hidCategory";
		upfrm.action = "/admin/ithinkso/category/category_item_process_ithinkso.asp";
		upfrm.mode.value="categoryitemdel";
		upfrm.submit();
		upfrm.CateDispSeqarr.value="";
		upfrm.target = "";
		upfrm.action = "";
		upfrm.mode.value="";
	}
	
</script>
			
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmSearch" method="post">
<input type="hidden" name="CateDispSeqarr">
<input type="hidden" name="iCateSeq0" value="<%=CateSeq0%>">
<input type="hidden" name="iCateSeq1" value="<%=CateSeq1%>">
<input type="hidden" name="iCateSeq2" value="<%=CateSeq2%>">
<input type="hidden" name="iCateSeq3" value="<%=CateSeq3%>">
<input type="hidden" name="Depth" value="<%=Depth%>">
<input type="hidden" name="mode">
<input type="hidden" name="menupos" value="<%= Request("menupos") %>">
<input type="hidden" name="reload" value="ON">
<input type="hidden" name="page">
<tr align="center" bgcolor="<%= adminColor("topbar") %>" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
		* �귣�� : <%	drawSelectBoxDesignerWithName "makerid", makerid %>
		&nbsp;&nbsp;
		* ��ǰ�ڵ� :
		<input type="text" class="text" name="itemid" value="<%= itemid %>" size="30" maxlength="100" onKeyPress="if (event.keyCode == 13) document.frm.submit();">(��ǥ�� �����Է°���)
		&nbsp;&nbsp;
		* ��ǰ�� :
		<input type="text" class="text" name="itemname" value="<%= itemname %>" size="32" maxlength="32">
		<p>
		* �Ǹſ���:
	   <select class="select" name="sellyn">
		   <option value="">��ü</option>
		   <option value="Y"  <%=CHKIIF(sellyn="Y","selected","")%>>�Ǹ�</option>
		   <option value="S"  <%=CHKIIF(sellyn="S","selected","")%>>�Ͻ�ǰ��</option>
		   <option value="N"  <%=CHKIIF(sellyn="N","selected","")%>>ǰ��</option>
		   <option value="YS"  <%=CHKIIF(sellyn="YS","selected","")%>>�Ǹ�+�Ͻ�ǰ��</option>
	   </select>
     	&nbsp;&nbsp;
     	* ī�װ���ǰ��뿩��:
	   <select class="select" name="usingyn">
		   <option value="">��ü</option>
		   <option value="Y"  <%=CHKIIF(usingyn="Y","selected","")%>>�����</option>
		   <option value="N"  <%=CHKIIF(usingyn="N","selected","")%>>������</option>
	   </select>		
	</td>
	
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="frmsubmit('');">
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("topbar") %>" >
	<td align="left">
		<br>
		<font color="Red">�� �űԵ�� �ϽǷ���, ī�װ��� �������ּ���.</font>
		<br>
		<%
		set cCate = new ccategory_ithinkso
			cCate.frectisusing = "Y"
			cCate.getCategoryType_notpaging
		%>
		* Ÿ�� :		
		<select name="CateSeq0" onchange="jsChCategory(0);">
			<option value="">--����--</option>
			<%
			IF cCate.fresultcount > 0 THEN
				
			For i = 0 To cCate.fresultcount - 1
			%>	
			<option value="<%= cCate.FItemList(i).fCateTypeSeq %>" <% if cstr(CateSeq0) = cstr(cCate.FItemList(i).fCateTypeSeq) then response.write " selected" %>>
				<%= cCate.FItemList(i).fCateTypeName %>
			</option>
			<%
			NEXT
			
			END IF	
			%>
		</select>
		<% set cCate = nothing %>
	 	&nbsp;>&nbsp;
		<% 
		set cCate = new ccategory_ithinkso
			cCate.frectCateTypeSeq = CateSeq0
			cCate.frectisusing = "Y"
			
			if CateSeq0 <> "" then
		 		cCate.getCategory_notpaging
		 	end if
		%>
		��ī�� : 
		<select name="CateSeq1" onChange="jsChCategory(1);">	
			<option value="">--��ü--</option>
			<%
			IF cCate.fresultcount > 0 THEN
				
			For i = 0 To cCate.fresultcount - 1
			%>				
			<option value="<%= cCate.FItemList(i).fCateSeq %>" <% if cstr(CateSeq1) = cstr(cCate.FItemList(i).fCateSeq) then response.write " selected" %>>
				<%= cCate.FItemList(i).fCateName %>
			</option>
			<%
			NEXT
			
			END IF	
			%>				
		</select>
		<% set cCate = nothing %>
		&nbsp;>&nbsp;		
		��ī�� :
		<% 
		set cCate = new ccategory_ithinkso
			cCate.frectCateTypeSeq = CateSeq0
			cCate.frectsubCateSeq1 = CateSeq1
			cCate.frectisusing = "Y"
			
			if CateSeq0 <> "" and Depth > 0 then
		 		cCate.getCategory_notpaging
		 	end if
		%>		
		<select name="CateSeq2" onChange="jsChCategory(2);">
			<option value="">--��ü--</option>
			<%
			IF cCate.fresultcount > 0 THEN
				
			For i = 0 To cCate.fresultcount - 1
			%>				
			<option value="<%= cCate.FItemList(i).fCateSeq %>" <% if cstr(CateSeq2) = cstr(cCate.FItemList(i).fCateSeq) then response.write " selected" %>>
				<%= cCate.FItemList(i).fCateName %>
			</option>
			<%
			NEXT
			
			END IF	
			%>
		</select>
		<% set cCate = nothing %>
		&nbsp;>&nbsp;		
		��ī�� :
		<% 
		set cCate = new ccategory_ithinkso
			cCate.frectCateTypeSeq = CateSeq0
			cCate.frectsubCateSeq1 = CateSeq1
			cCate.frectsubCateSeq2 = CateSeq2
			cCate.frectisusing = "Y"
			
			if CateSeq0 <> "" and Depth > 1 then
		 		cCate.getCategory_notpaging
		 	end if
		%>		
		<select name="CateSeq3" onChange="jsChCategory(3);">
			<option value="">--��ü--</option>
			<%
			IF cCate.fresultcount > 0 THEN
				
			For i = 0 To cCate.fresultcount - 1
			%>				
			<option value="<%= cCate.FItemList(i).fCateSeq %>" <% if cstr(CateSeq3) = cstr(cCate.FItemList(i).fCateSeq) then response.write " selected" %>>
				<%= cCate.FItemList(i).fCateName %>
			</option>
			<%
			NEXT
			
			END IF	
			%>
		</select>
		<% set cCate = nothing %>
		<input type="button" onclick="categoryitemreg();" class="button" value="ī�װ���ǰ���">
	</td>
</tr>
</form>
</table>

<br>
<!-- ǥ �߰��� ����-->
<table width="100%" align="center" cellpadding="1" cellspacing="1" class="a">	
<tr valign="bottom">       
    <td align="left">
    </td>
    <td align="right">
    	<input type="button" onclick="categoryitemdel(frmSearch);" class="button" value="���û���">
    </td>
</tr>	
</table>
<!-- ǥ �߰��� ��-->

<iframe id="hidCategory" name="hidCategory" src="about:blank" frameborder="0" width=0 height=0></iframe>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		<table width="100%" cellpadding="0" cellspacing="0" class="a">
		<tr>
			<td>
				�˻���� : <b><%= oitem.FTotalCount%></b>
				&nbsp;
				������ : <b><%= page %> /<%=  oitem.FTotalpage %></b>
			</td>
			<td align="right">
			</td>
		</tr>
		</table>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td><input type="checkbox" name="ckall" onclick="ckAll(this)"></td>
	<td>itemID</td>
	<td> �̹���</td>
	<td>�귣��ID</td>
	<td>��ǰ��</td>
	<td>ī�װ�</td>	
	<td>�ǸŰ�</td>
	<td>�Ǹ�<br>����</td>
	<td>ī�װ���ǰ<br>��뿩��</td>
	<td>����</td>
</tr>
<% if oitem.FresultCount<1 then %>
<tr bgcolor="#FFFFFF">
	<td colspan="15" align="center">[�˻������ �����ϴ�.]</td>
</tr>
<% end if %>
<% if oitem.FresultCount > 0 then %>
<% for i=0 to oitem.FresultCount-1 %>

<% if oitem.FItemList(i).Fisusing = "Y" then %>
	<tr align="center" bgcolor="#FFFFFF" onmouseover="this.style.background='#f1f1f1';" onmouseout="this.style.background='#FFFFFF';">
<% else %>
	<tr align="center" bgcolor="#e1e1e1" onmouseover="this.style.background='#f1f1f1';" onmouseout="this.style.background='#e1e1e1';">
<% end if %>

<form action="" name="frmBuyPrc<%=i%>" method="get">
	<input type="hidden" name="CateDispSeq" value="<%= oitem.FItemList(i).fCateDispSeq %>">
	<td align="center" width=30><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"></td>
	<td align="center" width=60>
		<input type="hidden" name="itemid" value="<%= oitem.FItemList(i).Fitemid %>">
		<a href="http://www.10x10.co.kr/shopping/category_prd.asp?itemid=<%= oitem.FItemList(i).Fitemid %>" target="_blank" title="�̸�����">				
		<%= oitem.FItemList(i).Fitemid %></a>
		</td>
	<td align="center" width=50><img src="<%= oitem.FItemList(i).FSmallImage %>" width="50" height="50" border="0"></td>
	<td align="left"><%= oitem.FItemList(i).Fmakerid %></td>
	<td align="left"><% =oitem.FItemList(i).Fitemname %></td>
	<td align="left">
		<% if oitem.FItemList(i).fCateTypename <> "" then %>
			[<%= oitem.FItemList(i).fCateTypename %>]
		<% end if %>
		<% if oitem.FItemList(i).fCatename1 <> "" then %>
			<Br><%= oitem.FItemList(i).fCatename1 %>
		<% end if %>
		<% if oitem.FItemList(i).fCatename2 <> "" then %>
			>> <%= oitem.FItemList(i).fCatename2 %>
		<% end if %>
		<% if oitem.FItemList(i).fCatename3 <> "" then %>
			>> <%= oitem.FItemList(i).fCatename3 %>
		<% end if %>		
	
	</td>
	<td align="right" width=80>
	<%
		Response.Write "" & FormatNumber(oitem.FItemList(i).Forgprice,0) & ""
		'���ΰ�
		if oitem.FItemList(i).Fsailyn="Y" then
			Response.Write "<br><font color=#F08050>(��)" & FormatNumber(oitem.FItemList(i).Fsailprice,0) & "</font>"
		end if
		'������
		if oitem.FItemList(i).FitemCouponYn="Y" then
			Select Case oitem.FItemList(i).FitemCouponType
				Case "1"
					'Response.Write "<br><font color=#5080F0>(��)" & FormatNumber(oitem.FItemList(i).Forgprice*((100-oitem.FItemList(i).FitemCouponValue)/100),0) & "</font>"
				Case "2"
					'Response.Write "<br><font color=#5080F0>(��)" & FormatNumber(oitem.FItemList(i).Forgprice-oitem.FItemList(i).FitemCouponValue,0) & "</font>"
			end Select
		end if
	%>
	</td>
	<td align="center" width=30><%= fnColor(oitem.FItemList(i).Fsellyn,"yn") %></td>
	<td align="center" width=80><%= fnColor(oitem.FItemList(i).Fisusing,"yn") %></td>
    <td align="center" width=30>
    </td>
</form>    
</tr>
<% next %>

<tr height="25" bgcolor="FFFFFF">
	<td colspan="15" align="center">
		<% if oitem.HasPreScroll then %>
		<a href="javascript:frmsubmit('<%= oitem.StartScrollPage-1 %>')">[pre]</a>
		<% else %>
			[pre]
		<% end if %>

		<% for i=0 + oitem.StartScrollPage to oitem.FScrollCount + oitem.StartScrollPage - 1 %>
			<% if i>oitem.FTotalpage then Exit for %>
			<% if CStr(page)=CStr(i) then %>
			<font color="red">[<%= i %>]</font>
			<% else %>
			<a href="javascript:frmsubmit('<%= i %>')">[<%= i %>]</a>
			<% end if %>
		<% next %>

		<% if oitem.HasNextScroll then %>
			<a href="javascript:frmsubmit('<%= i %>')">[next]</a>
		<% else %>
			[next]
		<% end if %>
	</td>
</tr>
<% end if %>
</table>

<% set oitem = nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
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
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stylepick/stylelifeCls.asp"-->

<%
Dim cd1,i,page,isusing ,oitem,deliverytype,sailyn,couponyn,menupos ,idx ,CD2
dim makerid,itemid , itemname,sellyn,danjongyn,limityn,mwdiv ,defaultmargin , SortMet
dim cdl ,cdm ,cds , overlap
	overlap = request("overlap")
	cdl = request("cdl")
	cdm = request("cdm")
	cds = request("cds")
	idx      = request("idx")
	cd1 = request("cd1")
	cd2 = request("cd2")
	SortMet = request("SortMet")			
	itemid      = request("itemid")
	itemname    = request("itemname")
	makerid     = request("makerid")
	sellyn      = request("sellyn")
	danjongyn   = request("danjongyn") 
	mwdiv       = request("mwdiv")
	limityn     = request("limityn") 
	sailyn      = request("sailyn")
	couponyn	= request("couponyn")
	defaultmargin = request("defaultmargin")
	deliverytype       = request("deliverytype")	
	menupos = request("menupos")
	page = request("page")
	if page = "" then page = 1
	isusing = "Y"
	if overlap = "" then overlap = "notoverlap"
		
'//��ǰ ����Ʈ
set oitem = new ClsStyleLife
	oitem.FPageSize = 50
	oitem.FCurrPage = page
	oitem.FRectCate_Large   = cdl
	oitem.FRectCate_Mid     = cdm
	oitem.FRectCate_Small   = cds
	oitem.FRectSortDiv      = SortMet
	oitem.FRectMakerid      = makerid
If itemid <> "" Then
	If IsNumeric(itemid) = "False" Then
		rw "<script>alert('��ǰ�ڵ�� ���ڸ� �Է��ϼ���');location.replace('/admin/stylepick/pop_evtitemAddInfo.asp');</script>"
	End If
End If	
	oitem.FRectItemid       = itemid
	oitem.FRectItemName     = itemname
	oitem.FRectSellYN       = sellyn
	oitem.FRectDanjongyn    = danjongyn
	oitem.FRectLimityn      = limityn
	oitem.FRectMWDiv        = mwdiv
	oitem.FRectDeliveryType = deliverytype
	oitem.FRectSailYn       = sailyn
	oitem.FRectCouponYn		= couponyn		
	oitem.frectcd1 = cd1
	oitem.frectcd2 = cd2
	oitem.frectisusing = isusing
	oitem.frectoverlap = overlap
	oitem.GetTmemeitemList()
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

function SelectItemsadd(){	
	var frm;
	var itemcount = 0;
	frm = document.frm;

	if(typeof(frm.chkitem) !="undefined"){
		if(!frm.chkitem.length){
			if(!frm.chkitem.checked){
				alert("������ ��ǰ�� �����ϴ�. ��ǰ�� ������ �ּ���");
				return;
			}
			frm.itemidarr.value = frm.chkitem.value;
			itemcount = 1;
		}else{
		
			for(i=0;i<frm.chkitem.length;i++){
				if(frm.chkitem[i].checked) {	   	    			
					if (frm.itemidarr.value==""){
						frm.itemidarr.value = frm.chkitem[i].value;				
					}else{
						frm.itemidarr.value = frm.itemidarr.value + "," +frm.chkitem[i].value;
					} 
					
				}	
				itemcount = frm.chkitem.length;
			}
			if (frm.itemidarr.value == ""){
				alert("������ ��ǰ�� �����ϴ�. ��ǰ�� ������ �ּ���");
				return;
			}
		}
	}else{
		alert("�߰��� ��ǰ�� �����ϴ�.");
		return;
	} 
	
	frm.action = "/admin/stylepick/stylelife_theme_process.asp";
	frm.mode.value = "evtitemadd";
	frm.target="view";
	frm.submit();
}

// �����Ȳ �˾�
function PopItemStock(itemid){
	var popwin = window.open("/admin/stock/itemcurrentstock.asp?menupos=709&itemid=" + itemid,"popitemstocklist","width=1000 height=600 scrollbars=yes resizable=yes");
	popwin.focus();
}

function jsSerach(){
	var frm;
	frm = document.frm;
	frm.target = "_self";
	frm.action ="pop_themeitemAddInfo.asp";
	frm.submit();
}

// ������ �̵�
function NextPage(ipage){
	document.frm.page.value= ipage;
	document.frm.target = "_self";
	document.frm.action ="pop_themeitemAddInfo.asp";
	document.frm.submit();
}

</script>

<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get">
<input type="hidden" name="page">
<input type="hidden" name="sType">
<input type="hidden" name="itemidarr">
<input type="hidden" name="itemcount" value="0">
<input type="hidden" name="mode">
<input type="hidden" name="idx" value="<%=idx%>">
<input type="hidden" name="defaultmargin" value="<%=defaultmargin%>">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="cd1" value="<%= cd1 %>">
<tr align="center" bgcolor="<%= adminColor("topbar") %>" >
	<td rowspan="2" width="30" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
	</td>	
	<td rowspan="2" width="30" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="javascript:jsSerach();">
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("topbar") %>" >
	<td align="left">
		�Ǹ�:<% drawSelectBoxSellYN "sellyn", sellyn %>     	      	
     	����:<% drawSelectBoxDanjongYN "danjongyn", danjongyn %>     	 
     	����:<% drawSelectBoxLimitYN "limityn", limityn %>     	 
     	���:<% drawSelectBoxMWU "mwdiv", mwdiv %>     	
     	����:<% drawSelectBoxSailYN "sailyn", sailyn %>
     	����:<% drawSelectBoxCouponYN "couponyn", couponyn %>     	
     	���:<% drawBeadalDiv "deliverytype",deliverytype %>
		<br>�귣�� :<%	drawSelectBoxDesignerWithName "makerid", makerid %>
		��ǰ�ڵ� :
		<input type="text" class="text" name="itemid" value="<%= itemid %>" size="40" maxlength="100" onKeyPress="if (event.keyCode == 13) document.frm.submit();">
		(��ǥ�� �����Է°���)
		<br>��ǰ�� :
		<input type="text" class="text" name="itemname" value="<%= itemname %>" size="32" maxlength="20">
		<!-- #include virtual="/common/module/categoryselectbox.asp"-->
		<br>
		�з�:<% Drawcategory "cd2",cd2," onchange='jsSerach();'","CD2" %>
		<!--
		<input type="radio" name="overlap" value="all" <% if overlap="all" then response.write " checked"%>>����ǰ
		<input type="radio" name="overlap" value="notoverlap" <% if overlap="notoverlap" then response.write " checked"%>>���Ͻ�Ÿ�Ͽ����������λ�ǰ����
		//-->
	</td>
</tr>    
</table>
	
<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">		
		<input type="button" value="���û�ǰ�߰�" onClick="SelectItemsadd()" class="button">
		<font color="red">�� "[ON]StyleLife>>StyleLife ����" ���� ���� �ش� ī�װ��� ��ǰ�� �����ž� ��ǰ�� ���Դϴ�.</font>
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
				�˻���� : <b><%= oitem.FTotalCount%></b>
				&nbsp;
				������ : <b><%= page %> /<%=  oitem.FTotalpage %></b>				
			</td>
			<td align="right">
				����:<% Drawsort "SortMet" ,SortMet ," onchange='jsSerach();'" %>				
			</td>			
		</tr>
		</table>
	</td>
	
</tr>
		
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td><input type="checkbox" name="chkAll" onClick="jsChkAll();"></td>
	<td>��Ÿ��</td>
	<td>�з�</td>	
	<td>��ǰID</td>
	<td>�̹���</td>
	<td>�귣��</td>
	<td>��ǰ��</td>
	<td>�ǸŰ�</td>
	<td>���԰�</td>
	<td nowrap>���<br>����</td>	
	<td nowrap>���<br>����</td>
	<td nowrap>�Ǹ�<br>����</td>	
	<td nowrap>���<br>����</td>	
	<td nowrap>����<br>����</td>	
	<td nowrap>���<br>��Ȳ</td>
</tr>
<% if oitem.FresultCount > 0 then %>
<% for i=0 to oitem.FresultCount-1 %>
<tr align="center" bgcolor="#FFFFFF" onmouseout="this.style.backgroundColor='#FFFFFF'" onmouseover="this.style.backgroundColor='#F1F1F1'">
	<td align="center">
		<input type="checkbox" name="chkitem" value="<%= oitem.FItemList(i).FItemid %>">
	</td>
	<td align="center">
		<%= oitem.FItemList(i).fcd1name %> (<%= oitem.FItemList(i).fcd1 %>)
	</td>
	<td align="center">
		<%= oitem.FItemList(i).fcd2name %> (<%= oitem.FItemList(i).fcd2 %>)
	</td>	
	<td align="center"><A href="http://www.10x10.co.kr/shopping/category_prd.asp?itemid=<%= oitem.FItemList(i).FItemId %>" target="_blank"><%= oitem.FItemList(i).FItemId %></a></td>
	<td align="center"><%IF oitem.FItemList(i).FSmallImage <> "" THEN%><img src="<%= oitem.FItemList(i).FSmallImage %>" width="50" height="50" border=0 alt=""><%END IF%></td>
		<td align="center"><% =oitem.FItemList(i).Fmakerid %></td>
	<td>&nbsp;<% =oitem.FItemList(i).Fitemname %></td>
	<td align="center">
		<%
		Response.Write FormatNumber(oitem.FItemList(i).Forgprice,0)
		'���ΰ�
		if oitem.FItemList(i).Fsailyn="Y" then
			Response.Write "<br><font color=#F08050>(��)" & FormatNumber(oitem.FItemList(i).Fsailprice,0) & "</font>"
		end if
		'������
		if oitem.FItemList(i).FitemCouponYn="Y" then
			Select Case oitem.FItemList(i).FitemCouponType
				Case "1"
					Response.Write "<br><font color=#5080F0>(��)" & FormatNumber(oitem.FItemList(i).Forgprice*((100-oitem.FItemList(i).FitemCouponValue)/100),0) & "</font>"
				Case "2"
					Response.Write "<br><font color=#5080F0>(��)" & FormatNumber(oitem.FItemList(i).Forgprice-oitem.FItemList(i).FitemCouponValue,0) & "</font>"
			end Select
		end if
		%>
	</td>
	<td align="center"><%
			Response.Write FormatNumber(oitem.FItemList(i).Forgsuplycash,0)
			'���ΰ�
			if oitem.FItemList(i).Fsailyn="Y" then
				Response.Write "<br><font color=#F08050>" & FormatNumber(oitem.FItemList(i).Fsailsuplycash,0) & "</font>"
			end if
		%></td>
	<td align="center"><%=fnColor(oitem.FItemList(i).IsUpcheBeasong(),"delivery")%></td>
	<td align="center"><%= fnColor(oitem.FItemList(i).Fmwdiv,"mw") %></td>
	<td align="center">
	<%= fnColor(oitem.FItemList(i).Fsellyn,"yn") %>
	</td>
	<td align="center">
	<%= fnColor(oitem.FItemList(i).Fisusing,"yn") %>
	</td>
	<td align="center"><%= fnColor(oitem.FItemList(i).Flimityn,"yn") %></td>
	<td align="center" nowrap>
		<a href="javascript:PopItemStock('<%= oitem.FItemList(i).FItemId %>')" title="�����Ȳ �˾�">[����]</a><br>
		<%IF oitem.FItemList(i).IsSoldOut() THEN%>
			<img src="http://webadmin.10x10.co.kr/images/soldout_s.gif" width="30" height="12">
		<%END IF%>
	</td>
</tr>
<% next %>
<tr>
	<td colspan="20" align="center" bgcolor="#FFFFFF">
	 	<% if oitem.HasPreScroll then %>
			<a href="javascript:NextPage('<%= oitem.StartScrollPage-1 %>')">[pre]</a>
		<% else %>
			[pre]
		<% end if %>
		<% for i=0 + oitem.StartScrollPage to oitem.FScrollCount + oitem.StartScrollPage - 1 %>
			<% if i>oitem.FTotalpage then Exit for %>
			<% if CStr(page)=CStr(i) then %>
			<font color="red">[<%= i %>]</font>
			<% else %>
			<a href="javascript:NextPage('<%= i %>')">[<%= i %>]</a>
			<% end if %>
		<% next %>
		<% if oitem.HasNextScroll then %>
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
<iframe id="view" name="view" width=300 width=300 frameborder=0 scrolling="no"></iframe>
<% set oitem = nothing %>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->


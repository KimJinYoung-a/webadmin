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
dim evtidx,title,subcopy,state,banner_img,startdate,enddate,isusing,regdate,comment
dim lastadminid,cd1,opendate,closedate,partMDid,partWDid ,oevent,SortMet
Dim catetype ,i,page ,oitem,deliverytype,sailyn,couponyn ,catename
dim makerid,itemid , itemname,sellyn,danjongyn,limityn,mwdiv ,defaultmargin
	SortMet = request("SortMet")
	evtidx = request("evtidx")
	menupos = request("menupos")	
	cd1 = request("cd1")
	catetype = request("catetype")	
	isusing = request("isusing")	
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
	if isusing = "" then isusing = "Y"
	
	if evtidx = "" then
		response.write "<script>"
		response.write "	alert('��ȹ����ȣ�� �����ϴ�');"
		response.write "	history.back();"
		response.write "</script>"
	end if
	
'//�̺�Ʈ����
set oevent = new cstylepick
	oevent.frectevtidx = evtidx
	
	if evtidx <> "" then
		oevent.fnGetEvent_item()
		
		if oevent.ftotalcount > 0 then			
			title = oevent.foneitem.ftitle
			subcopy = oevent.foneitem.fsubcopy
			state = oevent.foneitem.fstate
			banner_img = oevent.foneitem.fbanner_img
			startdate = left(oevent.foneitem.fstartdate,10)
			enddate = left(oevent.foneitem.fenddate,10)			
			regdate = oevent.foneitem.fregdate
			comment = oevent.foneitem.fcomment
			lastadminid = oevent.foneitem.flastadminid
			cd1 = oevent.foneitem.fcd1
			opendate = oevent.foneitem.fopendate
			closedate = oevent.foneitem.fclosedate
			partMDid = oevent.foneitem.fpartMDid
			partWDid = oevent.foneitem.fpartWDid
			catename = oevent.foneitem.fcatename
		else
			response.write "<script>"
			response.write "	alert('�ش� �̺�Ʈ�� ���� ���� �ʽ��ϴ�');"
			response.write "	history.back();"
			response.write "</script>"		
		end if	
	end if

set oevent = nothing
			
'//��ǰ ����Ʈ
set oitem = new cstylepick
	oitem.FPageSize = 50
	oitem.FCurrPage = page
	oitem.frectevtidx = evtidx
	oitem.FRectSortDiv      = SortMet
	oitem.FRectMakerid      = makerid
	oitem.FRectItemid       = itemid
	oitem.FRectItemName     = itemname
	oitem.FRectSellYN       = sellyn
	oitem.FRectDanjongyn    = danjongyn
	oitem.FRectLimityn      = limityn
	oitem.FRectMWDiv        = mwdiv
	oitem.FRectDeliveryType = deliverytype
	oitem.FRectSailYn       = sailyn
	oitem.FRectCouponYn		= couponyn
	oitem.frectisusing = isusing
	oitem.GetevtItemList()
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

function SelectItemsdel(){	
	var frm;
	var itemcount = 0;
	frm = document.frm;

	if(typeof(frm.chkitem) !="undefined"){
		if(!frm.chkitem.length){
			if(!frm.chkitem.checked){
				alert("������ ��ǰ�� �����ϴ�. ��ǰ�� ������ �ּ���");
				return;
			}
			frm.evtitemidxarr.value = frm.chkitem.value;
			itemcount = 1;
		}else{
		
			for(i=0;i<frm.chkitem.length;i++){
				if(frm.chkitem[i].checked) {	   	    			
					if (frm.evtitemidxarr.value==""){
						frm.evtitemidxarr.value = frm.chkitem[i].value;				
					}else{
						frm.evtitemidxarr.value = frm.evtitemidxarr.value + "," +frm.chkitem[i].value;
					} 
					
				}	
				itemcount = frm.chkitem.length;
			}
			if (frm.evtitemidxarr.value == ""){
				alert("������ ��ǰ�� �����ϴ�. ��ǰ�� ������ �ּ���");
				return;
			}
		}
	}else{
		alert("�߰��� ��ǰ�� �����ϴ�.");
		return;
	} 
	
	frm.action = "/admin/stylepick/stylepick_event_process.asp";
	frm.mode.value = "evtitemdel";
	frm.submit();
	frm.evtitemidxarr.value = "";
	frm.itemcount.value = "0";
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
	frm.action ="stylepick_event_item.asp";
	frm.submit();
}

// ������ �̵�
function NextPage(ipage){
	document.frm.page.value= ipage;
	document.frm.target = "_self";
	document.frm.action ="stylepick_event_item.asp";
	document.frm.submit();
}

//����ǰ�߰�
function addevtnewItem(evtidx,cd1){
	var addevtnewItem = window.open('/admin/stylepick/pop_evtitemAddInfo.asp?evtidx='+evtidx+'&cd1='+cd1+'&menupos=<%=menupos%>','addevtnewItem','width=1024,height=768,scrollbars=yes,resizable=yes');
	addevtnewItem.focus();
}

//������������
function backpage(evtidx){
	location.href="/admin/stylepick/stylepick_event.asp?menupos=<%=menupos%>";
}

</script>

<table width="100%" border="0" cellpadding="1" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="#FFFFFF" align="center">
	<td bgcolor="<%= adminColor("tabletop") %>">��ȹ����ȣ</td>
	<td><%= evtidx %></td>
	<td bgcolor="<%= adminColor("tabletop") %>">ī�װ�</td>
	<td><%= catename %></td>
	<td bgcolor="<%= adminColor("tabletop") %>">����</td>
	<td><%= title %></td>	
</tr>

	
<tr bgcolor="#FFFFFF" align="center">
	<td bgcolor="<%= adminColor("tabletop") %>">����</td>
	<td><%= geteventstate(state) %></td>
	<td bgcolor="<%= adminColor("tabletop") %>">����ī��</td>
	<td><%= subcopy %></td>	
	<td bgcolor="<%= adminColor("tabletop") %>">�Ⱓ</td>
	<td>
		������ : <%=startdate%>
		~ ������ : <%=enddate%>
   		<%
		if opendate <> "1900-01-01" and opendate <> "" then response.write " ����ó���� : " & opendate
		if closedate <> "1900-01-01" and closedate <> "" then response.write " ����ó���� : " & closedate
		%>
	</td>	
</tr>
</table>
<br>
<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="post">	
<input type="hidden" name="page" >
<input type="hidden" name="evtidx" value="<%=evtidx%>">
<input type="hidden" name="sType" >
<input type="hidden" name="evtitemidxarr">
<input type="hidden" name="itemcount" value="0">
<input type="hidden" name="mode">	
<input type="hidden" name="defaultmargin" value="<%=defaultmargin%>">
<input type="hidden" name="menupos" value="<%= menupos %>">
<tr align="center" bgcolor="<%= adminColor("topbar") %>" >
	<td rowspan="2" width="30" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">				
		��������:<% drawSelectBoxUsingYN "isusing", isusing %>
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
		��ǰ�� :
		<input type="text" class="text" name="itemname" value="<%= itemname %>" size="32" maxlength="20">					  	
	</td>
</tr>    
</table>
<br>	
<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
		<% if state <> "9" then %>
			<input type="button" value="��ǰ�߰�" onclick="addevtnewItem('<%=evtidx%>','<%=cd1%>');" class="button">
			<input type="button" value="���û�ǰ����" onClick="SelectItemsdel()" class="button">
			<input type="button" value="��ȹ���������" onclick="backpage();" class="button">
		<% end if %>
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
	<td>ī�װ�</td>	
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
<% if oitem.FItemList(i).fisusing = "Y" then %>
<tr align="center" bgcolor="#FFFFFF" onmouseover=this.style.background="orange"; onmouseout=this.style.background='ffffff';>
<% else %>    
<tr align="center" bgcolor="#FFFFaa" onmouseover=this.style.background="orange"; onmouseout=this.style.background='FFFFaa';>
<% end if %>
	<td align="center">
		<input type="checkbox" name="chkitem" value="<%= oitem.FItemList(i).fevtitemidx %>">
	</td>
	<td align="center">
		<%= oitem.FItemList(i).fcd1name %><%= oitem.FItemList(i).fcd2name %><%= oitem.FItemList(i).fcd3name %>
		[<%= oitem.FItemList(i).fcd1 %><%= oitem.FItemList(i).fcd2 %><%= oitem.FItemList(i).fcd3 %>]
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
	</td>webadmin.10x10.co.kr
	<td align="center"><%= fnColor(oitem.FItemList(i).Flimityn,"yn") %></td>
	<td align="center" nowrap>
		<a href="javascript:PopItemStock('<%= oitem.FItemList(i).FItemId %>')" title="�����Ȳ �˾�">[����]</a><br>
		<%IF oitem.FItemList(i).IsSoldOut() THEN%>
			<img src="http://scm.10x10.co.kr/images/soldout_s.gif" width="30" height="12">
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

<% set oitem = nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->


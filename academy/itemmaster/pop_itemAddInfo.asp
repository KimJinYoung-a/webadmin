<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' History : 2010.09.28 �ѿ�� ����
' Description : ��ǰ �߰� - ����, ����ǰ ��ǰ��Ͽ� ���
'				input - actionURL(db ó���� �ʿ��� �Ķ���ͱ��� ����) ex.acURL = "/admin/eventmanage/event/eventitem_process.asp?eC=1234"
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/academy/lib/academy_function.asp"-->
<!-- #include virtual="/academy/lib/classes/diyshopitem/DIYitemCls.asp"-->
<%
dim target, actionURL ,page ,cdl, cdm, cds ,i
dim itemid_s, itemname, makerid, sellyn, usingyn, deliverytype, limityn, vatyn, sailyn, couponyn, mwdiv,defaultmargin
	actionURL	= RequestCheckvar(request("acURL"),128)
	itemid_s      = request("itemid_s")
	itemname    = RequestCheckvar(request("itemname"),64)
	makerid     = RequestCheckvar(request("makerid"),32)
	sellyn      = RequestCheckvar(request("sellyn"),2)
	usingyn     = RequestCheckvar(request("usingyn"),2)	
	mwdiv       = RequestCheckvar(request("mwdiv"),2)
	limityn     = RequestCheckvar(request("limityn"),2)
	sailyn      = RequestCheckvar(request("sailyn"),2)
	couponyn	= RequestCheckvar(request("couponyn"),2)
	defaultmargin = RequestCheckvar(request("defaultmargin"),10)
	deliverytype       = RequestCheckvar(request("deliverytype"),2)
	cdl = RequestCheckvar(request("cdl"),10)
	cdm = RequestCheckvar(request("cdm"),10)
	cds = RequestCheckvar(request("cds"),10)
	page = RequestCheckvar(request("page"),10)
	
	if (page="") then page=1
	'if sailyn="" and instr(actionURL,"saleitem")>0 then sailyn="N"			'�������������� �˻��ȰŶ�� �⺻��: ���ξ���(������ ����)
	if couponyn="" and instr(actionURL,"saleitem")>0 then couponyn="N"
	'if sellyn = "" then sellyn ="Y"

	if itemid_s<>"" then
	dim iA ,arrTemp,arrItemid
	itemid_s = replace(itemid_s,",",chr(10))
	itemid_s = replace(itemid_s,chr(13),"")
	arrTemp = Split(itemid_s,chr(10))

	iA = 0
	do while iA <= ubound(arrTemp) 
		if trim(arrTemp(iA))<>"" then
			'��ǰ�ڵ� ��ȿ�� �˻�(2008.08.05;������)
			if Not(isNumeric(trim(arrTemp(iA)))) then
				Response.Write "<script language=javascript>alert('[" & arrTemp(iA) & "]��(��) ��ȿ�� ��ǰ�ڵ尡 �ƴմϴ�.');history.back();</script>"
				dbget.close()	:	response.End
			else
				arrItemid = arrItemid & trim(arrTemp(iA)) & ","
			end if
		end if
		iA = iA + 1
	loop
	itemid_s = left(arrItemid,len(arrItemid)-1)
	end if

dim oitem
set oitem = new CItem
	oitem.FPageSize         = 30
	oitem.FCurrPage         = page
	oitem.FRectMakerid      = makerid
	oitem.FRectItemid       = itemid_s
	oitem.FRectItemName     = itemname
	oitem.FRectSellYN       = sellyn
	oitem.FRectIsUsing      = usingyn	
	oitem.FRectLimityn      = limityn
	oitem.FRectMWDiv        = mwdiv
	oitem.FRectDeliveryType = deliverytype
	oitem.FRectsaleYn       = sailyn
	oitem.FRectCouponYn		= couponyn
	oitem.FRectCate_Large   = cdl
	oitem.FRectCate_Mid     = cdm
	oitem.FRectCate_Small   = cds
	oitem.GetItemList()		
%>

<script language="javascript">

function jsSerach(){
	var frm;
	frm = document.frm;
	frm.target = "_self";
	frm.action ="pop_itemAddInfo.asp";
	frm.submit();
}

function SelectItems(sType){	
var frm;
var itemcount = 0;
frm = document.frm;
frm.sType.value = sType;   //��ü���� or ���û�ǰ ���� ����

	if (sType == "sel"){
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
	   	    			 frm.itemidarr.value =  frm.chkitem[i].value;
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
	}else{
		if(typeof(frm.chkitem) !="undefined"){
			itemcount = "<%= oitem.FTotalCount%>";
		  if(confirm(itemcount +"���� �˻��� ��� ��ǰ�� �߰��Ͻðڽ��ϱ�?")){
		  	if(itemcount > 1000) {
		  		alert("��ǰ�� �ִ� 1000�Ǳ��� �����մϴ�. ������ �ٽ� �������ּ���");
		  		return;
		  	}
			frm.itemidarr.value = frm.itemid.value;
			
		  }else{
		  	return;
		  }
		}else{
		 	alert("�߰��� ��ǰ�� �����ϴ�.");
	   	  	return;
		}	
	}
	
	//frm.target = opener.name;
	frm.target = "FrameCKP";
	frm.action = "<%=actionURL%>";
	frm.itemcount.value = itemcount;
	frm.submit();
	frm.itemidarr.value = "";
	frm.itemcount.value = 0;	
	opener.location.reload(true);	
	//window.close();
}

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

// �����Ȳ �˾�
function PopItemStock(itemid){
	var popwin = window.open("/admin/stock/itemcurrentstock.asp?menupos=709&itemid=" + itemid,"popitemstocklist","width=1000 height=600 scrollbars=yes resizable=yes");
	popwin.focus();
}

// ������ �̵�
function NextPage(ipage){
	document.frm.page.value= ipage;
	document.frm.target = "_self";
	document.frm.action ="pop_itemAddInfo.asp";
	document.frm.submit();
}

</script>

<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="post">	
<input type="hidden" name="page" >
<input type="hidden" name="sType" >
<input type="hidden" name="itemidarr" >
<input type="hidden" name="itemcount" value="0">
<input type="hidden" name="mode" value="I">
<input type="hidden" name="acURL" value="<%=actionURL%>">
<input type="hidden" name="defaultmargin" value="<%=defaultmargin%>">
<tr align="center" bgcolor="<%= adminColor("topbar") %>" >
	<td rowspan="2" width="30" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
		<!-- include virtual="/common/module/categoryselectbox.asp"-->
		<br>
		�귣�� : <input type="text" class="text" name="makerid" value="<%= makerid %>" size="40" maxlength="100" onKeyPress="if (event.keyCode == 13) document.frm.submit();">	
		��ǰ�ڵ� :
		<textarea rows="3" cols="10" name="itemid_s" id="itemid_s"><%=replace(itemid_s,",",chr(10))%></textarea>			
		<br>��ǰ�� :
		<input type="text" class="text" name="itemname" value="<%= itemname %>" size="32" maxlength="20"> (��ǥ�� �����Է°���)
	</td>
	
	<td rowspan="2" width="30" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="javascript:jsSerach();">
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("topbar") %>" >
	<td align="left">
		�Ǹ�:<% drawSelectBoxSellYN "sellyn", sellyn %>
     	 
     	���:<% drawSelectBoxUsingYN "usingyn", usingyn %>
         	     	 
     	����:<% drawSelectBoxLimitYN "limityn", limityn %>
     	 
     	���:<% drawSelectBoxMWU "mwdiv", mwdiv %>
     	
     	����: <% drawSelectBoxSailYN "sailyn", sailyn %>

     	����: <% drawSelectBoxCouponYN "couponyn", couponyn %>
     	
     	<br>���:<% drawBeadalDiv "deliverytype",deliverytype %>
	</td>
</tr>    
</table>

<table width="100%" height="40" align="center" cellpadding="3" cellspacing="1" class="a" border="0">	
<tr>
	<td  valign="bottom">				
			<input type="button" value="���û�ǰ �߰�" onClick="SelectItems('sel')" class="button">
			<!-- saleItemProc.asp ��ü���� �߰� ���� �ִµ���. -->
			<input type="button" value="��ü���� �߰�" onClick="SelectItems('all')" class="button" disabled >
	</td>				
</tr>
</table>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>" valign="top" border="0">
<tr  bgcolor="#FFFFFF">
	<td colspan="13">
	�˻���� : <b><%= oitem.FTotalCount%></b>
	&nbsp;
	������ : <b><%= page %> /<%=  oitem.FTotalpage %></b>
	</td>		
</tr>
		
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td align="center"><input type="checkbox" name="chkAll" onClick="jsChkAll();"></td>
	<td align="center">��ǰID</td>
	<td align="center">�̹���</td>
	<td align="center">�귣��</td>
	<td align="center">��ǰ��</td>
	<td align="center">�ǸŰ�</td>
	<td align="center">���԰�</td>
	<td align="center" nowrap>���<br>����</td>	
	<td align="center" nowrap>���<br>����</td>
	<td align="center" nowrap>�Ǹ�<br>����</td>	
	<td align="center" nowrap>���<br>����</td>	
	<td align="center" nowrap>����<br>����</td>	
	<td align="center" nowrap>���<br>��Ȳ</td>
</tr>
<% if oitem.FresultCount<1 then %>
<tr bgcolor="#FFFFFF" >
	<td colspan="13" align="center">[�˻������ �����ϴ�.]</td>
</tr>
<% end if %>
<% if oitem.FresultCount > 0 then %>
<% for i=0 to oitem.FresultCount-1 %>
<tr class="a" height="25" bgcolor="#FFFFFF">
	<td  align="center"><input type="checkbox" name="chkitem" value="<%= oitem.FItemList(i).FItemId %>"></td>
	<td align="center"><A href="<%=wwwFingers%>/diyshop/shop_prd.asp?itemid=<%= oitem.FItemList(i).FItemId %>" target="_blank"><%= oitem.FItemList(i).FItemId %></a></td>
	<td align="center"><%IF oitem.FItemList(i).FSmallImage <> "" THEN%><img src="<%= oitem.FItemList(i).FSmallImage %>" width="50" height="50" border=0 alt=""><%END IF%></td>
	<td align="center"><% =oitem.FItemList(i).Fmakerid %></td>
	<td>&nbsp;<% =oitem.FItemList(i).Fitemname %></td>
	<td align="center">
		<%
		Response.Write FormatNumber(oitem.FItemList(i).Forgprice,0)
		'���ΰ�
		if oitem.FItemList(i).Fsaleyn="Y" then
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
	<td align="center">
		<%
		Response.Write FormatNumber(oitem.FItemList(i).Forgsuplycash,0)
		'���ΰ�
		if oitem.FItemList(i).Fsaleyn="Y" then
			Response.Write "<br><font color=#F08050>" & FormatNumber(oitem.FItemList(i).Fsailsuplycash,0) & "</font>"
		end if
		'������
		if oitem.FItemList(i).FitemCouponYn="Y" then
			if oitem.FItemList(i).FitemCouponType="1" or oitem.FItemList(i).FitemCouponType="2" then
				if oitem.FItemList(i).Fcouponbuyprice=0 or isNull(oitem.FItemList(i).Fcouponbuyprice) then
					Response.Write "<br><font color=#5080F0>" & FormatNumber(oitem.FItemList(i).Forgsuplycash,0) & "</font>"
				else
					Response.Write "<br><font color=#5080F0>" & FormatNumber(oitem.FItemList(i).Fcouponbuyprice,0) & "</font>"
				end if
			end if
		end if
		%>
	</td>
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
		<!--<a href="javascript:PopItemStock('<%= oitem.FItemList(i).FItemId %>')" title="�����Ȳ �˾�">[����]</a><br>-->
		<%IF oitem.FItemLiwebadmin.10x10.co.kr() THEN%>
			<img src="http://scm.10x10.co.kr/images/soldout_s.gif" width="30" height="12">
		<%END IF%>
	</td>
</tr>
<% next %>
<tr>
	<td colspan="13" align="center" bgcolor="#FFFFFF">
		<!-- ����¡ó�� -->
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
</form>
<% end if %>
</table>

<iframe name="FrameCKP" src="about:blank" frameborder="0" width="0" height="0"></iframe>

<%
 set oitem = nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->
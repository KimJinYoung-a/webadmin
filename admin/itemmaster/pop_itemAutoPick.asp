<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' History : 2017-06-27 ����ȭ ����
' Description : MD`PICK ���� ��ǰ
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbEVTopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/items/itemcls_2008.asp"-->
<%
dim actionURL ,page , gubun

actionURL 	= Replace(ReplaceRequestSpecialChar(request("acURL")),"||","&")
page = requestCheckvar(request("page"),10)
gubun = requestCheckvar(request("gubun"),1)

if (page="") then page=1

dim oitem , arrList
set oitem = new CItem
	oitem.FPageSize         = 30
	oitem.FCurrPage         = page
	oitem.Fgubun	        = gubun
	oitem.GetItemAutoPick
dim i

%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script language="javascript">
<!--
function SelectItems(sType){	
var frm;
var itemcount = 0;
frm = document.frmItem;
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
		  if(confirm("<%= oitem.FTotalCount%>���� �˻��� ��� ��ǰ�� �߰��Ͻðڽ��ϱ�?")){
		  	if(itemcount > 1000) {
		  		alert("��ǰ�� �ִ� 1000�Ǳ��� �����մϴ�. ������ �ٽ� �������ּ��� ");
		  		return;
		  	}
			frm.itemidarr.value = document.frm.itemid.value;
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
	opener.history.go(0);	
	//window.close();
}

function SelectAllItems(){	 
var frm;
var itemcount = 0;
frm = document.frm;  
		itemcount = "<%= oitem.FTotalCount%>"; 
		if (itemcount >0){
		  if(confirm("<%= oitem.FTotalCount%>���� �˻��� ��� ��ǰ�� �߰��Ͻðڽ��ϱ�?")){
		  	if(itemcount > 1000) {
		  		alert("��ǰ�� �ִ� 1000�Ǳ��� �����մϴ�. ������ �ٽ� �������ּ��� ");
		  		return;
		  	} 
		  }else{
		  	return;
		  }
		}else{
		 	alert("�߰��� ��ǰ�� �����ϴ�.");
	   	  	return;
		}	 
	 
	//frm.target = opener.name;
	frm.sType.value = "all";
	frm.target = "FrameCKP"; 
	frm.action = "<%=actionURL%>";
	frm.itemcount.value = itemcount;
	frm.submit();
	frm.itemidarr.value = "";
	frm.itemcount.value = 0;	
	opener.history.go(0);	
	//window.close();
}
//��ü ����
function jsChkAll(){	
var frm;
frm = document.frmItem;
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
	document.frm.action ="pop_itemAutoPick.asp";
	document.frm.submit();
}

//-->
</script>
<!-- �˻� ���� -->
<form name="frm" method="post">
	<input type="hidden" name="page" >
	<input type="hidden" name="sType" >
	<input type="hidden" name="itemidarr" >
	<input type="hidden" name="itemcount" value="0">
	<input type="hidden" name="mode" value="I">
	<input type="hidden" name="gubun" value="<%=gubun%>">
	<input type="hidden" name="acURL" value="<%=actionURL%>">
</form>
<form name="frmItem" method="post">	
	<input type="hidden" name="page" >
	<input type="hidden" name="sType" >
	<input type="hidden" name="itemidarr" >
	<input type="hidden" name="itemcount" value="0">
	<input type="hidden" name="mode" value="I">
	<input type="hidden" name="acURL" value="<%=actionURL%>">
<table width="100%" height="40" align="center" cellpadding="3" cellspacing="1" class="a" border="0">	
	<tr>
		<td  valign="bottom">				
			<input type="button" value="���û�ǰ �߰�" onClick="SelectItems('sel')" class="button">
		</td>				
	</tr>
</table> 
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>" valign="top" border="0">
<tr  bgcolor="#FFFFFF">
	<td colspan="14">
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
	<td align="center">������ȯ</td>
	<td align="center">Item<br>priority</td>
	<td align="center">�ֱٵ����</td>
	<td align="center">��۱���</td>
	<td align="center">�Ǹſ���</td>
	<td align="center">�����Ǹŷ�</td>	
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
	<td align="center"><A href="http://www.10x10.co.kr/shopping/category_prd.asp?itemid=<%= oitem.FItemList(i).FItemId %>" target="_blank"><%= oitem.FItemList(i).FItemId %></a></td>
	<td align="center"><%IF oitem.FItemList(i).FSmallImage <> "" THEN%><img src="<%= oitem.FItemList(i).FSmallImage %>" width="50" height="50" border=0 alt=""><%END IF%></td>
	<td align="center"><% =oitem.FItemList(i).Fmakerid %></td>
	<td align="left"><font color="gray">(<%=oitem.FItemList(i).Fcatename%>)</font><br/><%=oitem.FItemList(i).Fitemname %></td>
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
	<td align="center">
		<%
			Response.Write FormatNumber(oitem.FItemList(i).Forgsuplycash,0)
			'���ΰ�
			if oitem.FItemList(i).Fsailyn="Y" then
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
	<td align="center"><%=oitem.FItemList(i).Forderedcnt%><br/>(<%=CInt(oitem.FItemList(i).Fcr)%>)%</td>
	<td align="center"><%=oitem.FItemList(i).Ftotalwgt%></td>
	<td align="center">
	<% If isnull(oitem.FItemList(i).Flastregdt) Then 
		Response.write "-"
	   Else
		If oitem.FItemList(i).Flastregdt=0 Then 
			Response.write "�Ǹ���"
		Else 
			Response.write Replace(oitem.FItemList(i).Flastregdt,"-","")&"����"
		End If 
	   End If 
	%>
	</td>
	<td align="center"><%=fnColor(oitem.FItemList(i).IsUpcheBeasong(),"delivery")%></td>
	<td align="center"><%=oitem.FItemList(i).Fsellyn%></td>
	<td align="center"><%=oitem.FItemList(i).Fyesterdaysales%></td>
	<td align="center" nowrap>
	<a href="javascript:PopItemStock('<%= oitem.FItemList(i).FItemId %>')" title="�����Ȳ �˾�">[����]</a><br>
	<%IF oitem.FItemList(i).IsSoldOut() THEN%>
		<img src="http://webadmin.10x10.co.kr/images/soldout_s.gif" width="30" height="12">
<%END IF%>
	</td>
</tr>
<% next %>
<% end if %>
</table>
</form>
<table width="100%"   align="center" cellpadding="3" cellspacing="1" class="a" border="0">	
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
</table> 

<div style="padding:5px;text-align:right;font-size:8pt">Ver1.0  lastupdate: 2017-06-27 </div>
<iframe name="FrameCKP" src="about:blank" frameborder="0" width="600" height="200"></iframe>
<%
 set oitem = nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbEVTclose.asp" -->
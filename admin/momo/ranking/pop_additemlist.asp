<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/items/itemcls_2008.asp"-->
<%
dim target,gubun
dim eCode, cEGroup,arrGroup,intLoop
dim itemid, itemname, makerid, sellyn, usingyn, danjongyn, mwdiv, limityn, vatyn, sailyn,deliverytype
dim cdl, cdm, cds
dim page
Dim sortDiv

eCode =request("eC")
itemid      = request("itemid")
itemname    = request("itemname")
makerid     = request("makerid")
sellyn      = request("sellyn")
usingyn     = request("usingyn")
danjongyn   = request("danjongyn") 
limityn     = request("limityn") 
sailyn      = request("sailyn")
mwdiv       = request("mwdiv")
deliverytype       = request("deliverytype")
sortDiv 	= request("sortDiv")

cdl = request("cdl")
cdm = request("cdm")
cds = request("cds")

page = request("page")

if (page="") then page=1
'if sellyn = "" then sellyn ="Y"
if itemid<>"" then
	dim iA ,arrTemp,arrItemid

	arrTemp = Split(itemid,",")

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

	itemid = left(arrItemid,len(arrItemid)-1)
end if


	if sortDiv="" then sortDiv="new"	'���Ĺ�� �⺻��


'==============================================================================
dim oitem

set oitem = new CItem

oitem.FPageSize         = 30
oitem.FCurrPage         = page
oitem.FRectMakerid      = makerid
oitem.FRectItemid       = itemid
oitem.FRectItemName     = itemname

oitem.FRectSellYN       = sellyn
oitem.FRectIsUsing      = usingyn
oitem.FRectDanjongyn    = danjongyn
oitem.FRectLimityn      = limityn
oitem.FRectMWDiv        = mwdiv
oitem.FRectSailYn       = sailyn
oitem.FRectDeliveryType = deliverytype

oitem.FRectCate_Large   = cdl
oitem.FRectCate_Mid     = cdm
oitem.FRectCate_Small   = cds
oitem.FRectSortDiv = SortDiv

oitem.GetItemList

dim i

			
%>
<script language="javascript">
<!--
document.domain = "10x10.co.kr";

function jsSerach(){
	//var frm;
	//frm = document.frm;
	//frm.target = "_self";
	//frm.action = "/admin/momo/ranking/pop_additemlist.asp";
	document.frm.submit();
}

function SelectItems(sType){
var itemcount = 0;
var frm;
frm = document.frm;
frm.sType.value = sType;

	frm.itemidarr.value = "";
	frm.itemnamearr.value = "";

	if (sType == "sel"){
		 if(typeof(frm.chkitem) !="undefined"){
	   	   	if(!frm.chkitem.length){
	   	   		if(!frm.chkitem.checked){
	   	   			alert("������ ��ǰ�� �����ϴ�. ��ǰ�� ������ �ּ���");
	   	   			return;
	   	   		}
	   	   		 frm.itemidarr.value = frm.chkitem.value;
	   	   		 frm.itemnamearr.value = frm.chkitemname.value;
	   	    }else{
	   	    	for(i=0;i<frm.chkitem.length;i++){
	   	    		if(frm.chkitem[i].checked) {
	   	    			if (frm.itemidarr.value==""){
	   	    			 frm.itemidarr.value =  frm.chkitem[i].value;
	   	    			 frm.itemnamearr.value =  frm.chkitemname[i].value;
	   	    			}else{
	   	    			 frm.itemidarr.value = frm.itemidarr.value + "," +frm.chkitem[i].value;
	   	    			 frm.itemnamearr.value = frm.itemnamearr.value + "," +frm.chkitemname[i].value;
	   	    			} 
	   	    		}	
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

	opener.document.frm.itemid.value = frm.itemidarr.value;
	//opener.document.frm.itemname.value = frm.itemnamearr.value;
	window.close();
}

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
	document.frm.submit();
}

//-->
</script>
<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="post" action="/admin/momo/ranking/pop_additemlist.asp">
	<input type="hidden" name="eC" value="<%=eCode%>">
	<input type="hidden" name="page" >
	<input type="hidden" name="sType" >
	<input type="hidden" name="itemidarr" >
	<input type="hidden" name="itemnamearr" >
	<input type="hidden" name="mode" value="I">
	<tr align="center" bgcolor="<%= adminColor("topbar") %>" >
		<td rowspan="2" width="30" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
		<td align="left">
			�귣�� :<%	drawSelectBoxDesignerWithName "makerid", makerid %>
			&nbsp;
			<!-- #include virtual="/common/module/categoryselectbox.asp"-->
			<br>
			��ǰ�ڵ� :
			<input type="text" class="text" name="itemid" value="<%= itemid %>" size="40" maxlength="100" onKeyPress="if (event.keyCode == 13) document.frm.submit();">
			&nbsp;
			��ǰ�� :
			<input type="text" class="text" name="itemname" value="<%= itemname %>" size="32" maxlength="20">
			<div style="font-size:11px; color:gray;padding-left:60px;">(��ǥ�� �����Է°���)</div>
		</td>
		
		<td rowspan="2" width="30" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="�˻�" onClick="javascript:jsSerach();">
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("topbar") %>" >
		<td align="left">
			�Ǹ�:<% drawSelectBoxSellYN "sellyn", sellyn %>
	     	 
	     	���:<% drawSelectBoxUsingYN "usingyn", usingyn %>
	         	
	     	����:<% drawSelectBoxDanjongYN "danjongyn", danjongyn %>
	     	 
	     	����:<% drawSelectBoxLimitYN "limityn", limityn %>

	     	���� <% drawSelectBoxSailYN "sailyn", sailyn %>
	     	
	     	<br>���:<% drawBeadalDiv "deliverytype",deliverytype %>
		</td>
	</tr>    
</table>

<table width="100%" height="40" align="center" cellpadding="3" cellspacing="1" class="a" border="0">
	<tr>
		<td  valign="bottom">
			<input type="button" value="���û�ǰ �߰�" onClick="SelectItems('sel')" class="button">
		</td>
	</tr>
</table>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr  bgcolor="#FFFFFF">
	<td colspan="9">
	�˻���� : <b><%= oitem.FTotalCount%></b>
	&nbsp;
	������ : <b><%= page %> /<%=  oitem.FTotalpage %></b>
	</td>
	<td colspan="3">
		<select name="sortDiv" onchange="this.form.submit();">
		<option value="new" <% IF sortDiv="new" Then response.write "selected" %> >�Ż�ǰ��</option>
		<option value="cashH" <% IF sortDiv="cashH" Then response.write "selected" %>>�������ݼ�</option>
		<option value="cashL" <% IF sortDiv="cashL" Then response.write "selected" %>>�������ݼ�</option>
		<option value="best" <% IF sortDiv="best" Then response.write "selected" %>>����Ʈ��</option>
		</select>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td align="center"></td>
	<td align="center">��ǰID</td>
	<td align="center">�̹���</td>
	<td align="center">�귣��</td>
	<td align="center">��ǰ��</td>
	<td align="center">�ǸŰ�</td>
	<td align="center">���԰�</td>
	<td align="center">���</td>
	<td align="center">�Ǹſ���</td>
	<td align="center">��뿩��</td>
	<td align="center">��������</td>
	<td align="center">�����Ȳ</td>
</tr>
<% if oitem.FresultCount<1 then %>
    <tr bgcolor="#FFFFFF">
    	<td colspan="12" align="center">[�˻������ �����ϴ�.]</td>
    </tr>
<% end if %>
<% if oitem.FresultCount > 0 then %>
    <% for i=0 to oitem.FresultCount-1 %>
	<tr class="a" height="25" bgcolor="#FFFFFF">
	<td  align="center"><input type="radio" name="chkitem" value="<%= oitem.FItemList(i).FItemId %>"><input type="hidden" name="chkitemname" value="<% =oitem.FItemList(i).Fitemname %>"></td>
	<td align="center"><A href="http://www.10x10.co.kr/shopping/category_prd.asp?itemid=<%= oitem.FItemList(i).FItemId %>" target="_blank"><%= oitem.FItemList(i).FItemId %></a></td>
	<td align="center"><%IF oitem.FItemList(i).FSmallImage <> "" THEN%><img src="<%= oitem.FItemList(i).FSmallImage %>" width="50" height="50" border=0 alt=""><%END IF%></td>
		<td align="center"><% =oitem.FItemList(i).Fmakerid %></td>
	<td>&nbsp;<% =oitem.FItemList(i).Fitemname %></td>
	<td align="center"><%
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
		%></td>
	<td align="center"><%
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
		%></td>
	<td align="center"><%=fnColor(oitem.FItemList(i).IsUpcheBeasong(),"delivery")%></td>
	<td align="center">
	<%= fnColor(oitem.FItemList(i).Fsellyn,"yn") %>
	</td>
	<td align="center">
	<%= fnColor(oitem.FItemList(i).Fisusing,"yn") %>
	</td>
	<td align="center"><%= fnColor(oitem.FItemList(i).Flimityn,"yn") %></td>
	<td align="center">
	<a href="javascript:PopItemStock('<%= oitem.FItemList(i).FItemId %>')" title="�����Ȳ �˾�">[����]</a><br>
	<%IF oitem.FItemList(i).IsSoldOut() THEN%>
		<img src="http://webadmin.10x10.co.kr/images/soldout_s.gif" width="30" height="12">
<%END IF%>
	</td>
</tr>
<% next %>
<tr>
	<td colspan="12" align="center" bgcolor="#FFFFFF">
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
</table>
<% end if %>
<iframe name="FrameCKP" src="" frameborder="0" width="0" height="0"></iframe>
<%
 set oitem = nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
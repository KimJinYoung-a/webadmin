<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/items/itemcls_2008.asp"-->
<!-- #include virtual="/lib/classes/items/dealManageCls.asp"-->
<%
dim target,gubun
dim idx, cEGroup,arrGroup,intLoop
dim itemid, itemname, makerid, sellyn, usingyn, danjongyn, mwdiv, limityn, vatyn, sailyn,deliverytype, keyword,couponyn, itemdiv
dim cdl, cdm, cds , dispCate
dim page
Dim sortDiv, notdeal
dim eChannel, pagesize, groupSelect

idx = requestCheckvar(request("idx"),9)
itemid      = request("itemid")
itemname    = requestCheckvar(request("itemname"),64)
makerid     = requestCheckvar(request("makerid"),32)
sellyn      = requestCheckvar(request("sellyn"),2)
usingyn     = requestCheckvar(request("usingyn"),1)
danjongyn   = requestCheckvar(request("danjongyn"),2) 
limityn     = requestCheckvar(request("limityn"),2) 
sailyn      = requestCheckvar(request("sailyn"),1)  
deliverytype= requestCheckvar(request("deliverytype"),1)
sortDiv 		= requestCheckvar(request("sortDiv"),5)
keyword			= requestCheckvar(request("keyword"),512)
eChannel    = requestCheckvar(request("eCh"),1)   
couponyn		= requestCheckvar(request("couponyn"),1)
itemdiv		= requestCheckvar(request("itemdiv"),2)
groupSelect		= requestCheckvar(request("groupSelect"),10)

cdl = requestCheckvar(request("cdl"),10)
cdm = requestCheckvar(request("cdm"),10)
cds = requestCheckvar(request("cds"),10)
dispCate = requestCheckvar(request("disp"),16)

page = requestCheckvar(request("page"),10)
notdeal = requestCheckvar(request("notdeal"),1)
if itemid<>"" then
	pagesize=680
else
	pagesize=105
end if
if (page="") then page=1
'if sellyn = "" then sellyn ="Y"
if itemid<>"" then
	dim iA ,arrTemp,arrItemid
	itemid = replace(itemid,chr(13),"")
	arrTemp = Split(itemid,chr(10))

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

oitem.FPageSize         = pagesize
oitem.FCurrPage         = page
oitem.FRectMakerid      = makerid
oitem.FRectItemid       = itemid
oitem.FRectItemName     = itemname
oitem.FRectKeyword		= keyword

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
oitem.FRectDispCate		= dispCate
oitem.FRectSortDiv = SortDiv
oitem.FRectCouponYn		= couponyn
oitem.FRectItemDiv		= itemdiv
oitem.FRectDealYn="Y"
oitem.GetItemList

dim i, cdealGroup, arrP, intP

set cdealGroup = new CDealSelect
cdealGroup.FRectDealCode = idx
arrP = cdealGroup.fnGetRootGroup
set cdealGroup = nothing
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script language="javascript">
<!--
function jsSerach(){
	var frm;
	frm = document.frm;
	frm.target = "_self";
	frm.action ="pop_deal_additemlist.asp";
	frm.submit();
}

function SelectItems(sType){	
var itemcount = 0;
var frm;
var ck=0;
frm = document.frm;
frm.sType.value = sType;

	if(sType == "sel"){
		if(typeof(frm.chkitem) !="undefined"){
	   	   	if(!frm.chkitem.length){
	   	   		if(!frm.chkitem.checked){
	   	   			alert("������ ��ǰ�� �����ϴ�. ��ǰ�� ������ �ּ���");
	   	   			return;
	   	   		}
	   	   		frm.itemidarr.value = frm.chkitem.value;
				frm.sitemnamearr.value = frm.sitemname.value;
	   	    }else{
	   	    	for(i=0;i<frm.chkitem.length;i++){
	   	    		if(frm.chkitem[i].checked) {
						ck=ck+1;	   	    			
	   	    			if (frm.itemidarr.value==""){
							frm.itemidarr.value =  frm.chkitem[i].value;
							frm.sitemnamearr.value = frm.sitemname[i].value;
	   	    			}else{
							frm.itemidarr.value = frm.itemidarr.value + "," + frm.chkitem[i].value;
							frm.sitemnamearr.value = frm.sitemnamearr.value + "|" + frm.sitemname[i].value;
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
		//alert(frm.chkitem.length);
		for(i=0;i<frm.chkitem.length;i++){
			frm.chkitem[i].checked = true;
			if (frm.itemidarr.value==""){
				frm.itemidarr.value =  frm.chkitem[i].value;
				frm.sitemnamearr.value = frm.sitemname[i].value;
			}else{
				frm.itemidarr.value = frm.itemidarr.value + "," +frm.chkitem[i].value;
				frm.sitemnamearr.value = frm.sitemnamearr.value + "|" + frm.sitemname[i].value;
			}
		}
	}
	console.log(frm.itemidarr.value);
	console.log(frm.sitemnamearr.value);
	$.ajax({
		type: "POST",
		url: "doDealItemSet.asp",
		data: "mode=add&idx=<%=idx%>&group_code="+frm.group_code.value+"&itemidarr="+frm.itemidarr.value+"&sitemnamearr="+escape(frm.sitemnamearr.value),
		contentType: 'application/x-www-form-urlencoded; charset=euc-kr',
		cache: false,
		success: function(message) {
			if(message=="err1"){
				alert("������ ������ �����ϴ�.");
			}else if(message=="err2"){
				alert("��ȿ���� ���� ���ڰ� ���ԵǾ� �ֽ��ϴ�.");
			}else if(message=="err3"){
				alert("������ ó���� ������ �߻��Ͽ����ϴ�.");
			}else{
				opener.location.reload();
				opener.jsItemAddLoad();
			}
		},
		error: function(err) {
			alert(err.responseText);
		}
	});
	frm.itemidarr.value = "";
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
	document.frm.action ="pop_deal_additemlist.asp";
	document.frm.submit();
}

//-->
</script>
<!-- �˻� ���� -->
<form name="frm" method=post>
<input type="hidden" name="idx" value="<%=idx%>">
<input type="hidden" name="page">
<input type="hidden" name="sType">
<input type="hidden" name="itemidarr">
<input type="hidden" name="sitemnamearr">
<input type="hidden" name="mode" value="I">
<input type="hidden" name="eCh" value="<%=eChannel%>">
<input type="hidden" name="groupSelect" value="<%=groupSelect%>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" border="0" > 	
<tr>
    <td>    
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>"> 
	<tr align="center" bgcolor="<%= adminColor("topbar") %>" >
		<td rowspan="2" width="60" bgcolor="<%= adminColor("gray") %>">�˻� ����</td>
		<td align="left">
			<table border="0" cellpadding="1" cellspacing="0" class="a">
				<tr>
					<td style="white-space:nowrap;">�귣��: <%	drawSelectBoxDesignerWithName "makerid", makerid %></td> 
					<td style="white-space:nowrap;padding-left:5px;">��ǰ��: <input type="text" class="text" name="itemname" value="<%= itemname %>" size="32" maxlength="20"></td>
					<td style="white-space:nowrap;padding-left:5px;">��ǰ�ڵ�:</td>
					<td style="white-space:nowrap;" rowspan="2"><textarea rows="3" cols="10" name="itemid" id="itemid"><%=replace(itemid,",",chr(10))%></textarea> </td>
				</tr>	 
			  <tr>	
			  	<td style="white-space:nowrap;"> <!-- #include virtual="/common/module/categoryselectbox.asp"--></td>
			    <td style="white-space:nowrap;padding-left:5px;" colspan="2">����ī�װ� : <!-- #include virtual="/common/module/dispCateSelectBox.asp"--></td>
			  </tr>   
	 		<tr>
	 			<td colspan="4">�˻�Ű���� : <input type="text" class="text" name="keyword" value="<%=keyword%>" size="40"><font color="gray" size="2">(����:�������ֽ��ϴ�.)</font></td>
	 		</tr>
	 	</table>
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
	     	
	     	����: <% drawSelectBoxCouponYN "couponyn", couponyn %>
	     	
			���:<% drawBeadalDiv "deliverytype",deliverytype %>

			��ǰ����:<% drawSelectBoxItemDivDeal "itemdiv",itemdiv %>

			<input type="checkbox" name="notdeal" value="Y"<% if notdeal="Y" then response.write " checked" %>> ����ǰ ����
		</td>
	</tr>    
</table>
</td>
</tr>
<tr>
    <td>
		<table width="100%" height="40" align="center" cellpadding="3" cellspacing="1" class="a" border="0">	
			<tr>
				<td valign="bottom">
					<%IF isArray(arrP) THEN %>
						<select name="group_code">
						<% For intP=0 To UBound(arrP,2) %>
							<option value="<%=arrP(0,intP)%>"<% IF Cstr(groupSelect) = Cstr(arrP(0,intP)) THEN %> selected<% END IF %>><%=arrP(1,intP)%></option>
						<% Next %>
						</select>
					<% else %>
					<input type="hidden" value="0" name="group_code">
					<% END IF %>
					<input type="button" value="���û�ǰ �߰�" onClick="SelectItems('sel')" class="button">
					<input type="button" value="��ü���� �߰�" onClick="SelectItems('all')" class="button">
				</td>				
			</tr>
		</table>
	</td>
</tr>
<tr>
    <td>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr  bgcolor="#FFFFFF">
	<td colspan="9">
	�˻���� : <b><%= oitem.FTotalCount%></b>
	&nbsp;
	������ : <b><%= page %> /<%=  oitem.FTotalpage %></b>
	</td>
	<td colspan="3">
		<select name="sortDiv" onchange="this.form.submit();" class="select">
		<option value="new" <% IF sortDiv="new" Then response.write "selected" %> >�Ż�ǰ��</option>
		<option value="cashH" <% IF sortDiv="cashH" Then response.write "selected" %>>�������ݼ�</option>
		<option value="cashL" <% IF sortDiv="cashL" Then response.write "selected" %>>�������ݼ�</option>
		<option value="best" <% IF sortDiv="best" Then response.write "selected" %>>����Ʈ��</option>
		</select>
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
	<td  align="center"><input type="checkbox" name="chkitem" value="<%= oitem.FItemList(i).FItemId %>"><input type='hidden' name='sitemname' id='sitemname' value='<%= oitem.FItemList(i).Fitemname %>'></td>
	<td align="center"><A href="http://www.10x10.co.kr/shopping/category_prd.asp?itemid=<%= oitem.FItemList(i).FItemId %>" target="_blank"><%= oitem.FItemList(i).FItemId %></a></td>
	<td align="center"><%IF oitem.FItemList(i).FSmallImage <> "" THEN%><img src="<%= oitem.FItemList(i).FSmallImage %>" width="50" height="50" border=0 alt=""><%END IF%></td>
		<td align="center"><% =oitem.FItemList(i).Fmakerid %></td>
	<td><%If oitem.FItemList(i).Fitemdiv="21" Then%><font color="#0000ff">[��]</font><%End If%>&nbsp;<% =oitem.FItemList(i).Fitemname %></td>
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
</table>
</td>
</tr>
</table> 
</form>
<% end if %>
<iframe name="FrameCKP" src="" frameborder="0" width="600" height="400"></iframe>
<%
 set oitem = nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
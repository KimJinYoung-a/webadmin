<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/items/itemcls_2008.asp"-->
<!-- #include virtual="/lib/classes/event/eventManageCls_V3.asp"-->
<%
dim target,gubun
dim eCode, cEGroup,arrGroup,intLoop,egcode
dim itemid, itemname, makerid, sellyn, usingyn, danjongyn, mwdiv, limityn, vatyn, sailyn,deliverytype, keyword,couponyn, itemdiv
dim cdl, cdm, cds , dispCate
dim page
Dim sortDiv
dim eChannel
dim formName

formName = request("formName")
eCode 			= requestCheckvar(request("eC"),10)
egcode		= requestCheckvar(request("egcode"),10)
itemid      = requestCheckvar(request("itemid"),255)
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

cdl = requestCheckvar(request("cdl"),10)
cdm = requestCheckvar(request("cdm"),10)
cds = requestCheckvar(request("cds"),10)
dispCate = requestCheckvar(request("disp"),16)

page = requestCheckvar(request("page"),10)

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

'--�̺�Ʈ �׷�
 set cEGroup = new ClsEventGroup
	cEGroup.FECode = eCode 
	cEGroup.FEChannel = eChannel	
	arrGroup = cEGroup.fnGetEventItemGroup		
 set cEGroup = nothing


'==============================================================================
dim oitem

set oitem = new CItem

oitem.FPageSize         = 30
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
oitem.FRectDealYn		= "N"
oitem.GetItemList

dim i

			
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script language="javascript">
<!--
function jsSerach(){
	var frm;
	frm = document.frm;
	frm.target = "_self";
	frm.action ="item_regist.asp";
	frm.submit();
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
	document.frm.action ="item_regist.asp";
	document.frm.submit();
}
function addItem(itemid, itemimg, itemname, saleper){		
	var frm = window.opener.document.frm;		
	if(confirm("��ǰ�� �߰��Ͻðڽ��ϱ�?")){		
		frm.itemid1.value=itemid;		
		frm.evttitle.value=itemname;			
		frm.evtimg.value=itemimg
		window.opener.document.getElementById("evtimgsrc").src = itemimg		
	}
	window.close();
}
$(function(){
	$(".a tr td a").click(function(e){
		e.stopPropagation();
	});		
});
//-->
</script>
<!-- �˻� ���� -->
<form name="frm" method=get>	
	<input type="hidden" name="eC" value="<%=eCode%>">
	<input type="hidden" name="page" >
	<input type="hidden" name="sType" >
	<input type="hidden" name="itemidarr" >
	<input type="hidden" name="mode" value="I">
	<input type="hidden" name="eCh" value="<%=eChannel%>">
	<input type="hidden" name="egcode" value="<%=egcode%>">
	<input type="hidden" name="formName" value="<%=formName%>">
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
		</td>
	</tr>    
</table>
</td>
</tr>
<tr>
    <td>
<table width="100%" align="center"    cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
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
<% 	dim tempSalePer
	if oitem.FresultCount > 0 then 	
%>
    <% 
		dim itemImg		
		for i=0 to oitem.FresultCount-1 
		tempSalePer = CLng((oitem.FItemList(i).Forgprice-oitem.FItemList(i).Fsellcash)/oitem.FItemList(i).FOrgPrice*100)

		if oitem.FItemList(i).Ftentenimage600 <> "" then
			itemimg = oitem.FItemList(i).Ftentenimage600
		else
			if oitem.FItemList(i).Fitemdiv = 21 then
				itemImg = getThumbImgFromURL(oitem.FItemList(i).Fbasicimage,600,600,"true","false")
			else
				itemimg = oitem.FItemList(i).Fbasicimage
			end if
		end if
	%>
	
	<tr class="a" onclick="addItem('<%=oitem.FItemList(i).FItemId%>','<%=itemImg%>','<% =oitem.FItemList(i).Fitemname %>', '<%=chkIIF( tempSalePer = 0, "", tempSalePer&"%")%>');" height="25" bgcolor="#FFFFFF" style="cursor:pointer;" onmouseover="this.style.backgroundColor='#D8D8D8'" onmouseout="this.style.backgroundColor=''">	
	<td align="center"><A href="http://www.10x10.co.kr/shopping/category_prd.asp?itemid=<%= oitem.FItemList(i).FItemId %>" target="_blank"><%= oitem.FItemList(i).FItemId %></a></td>
	<td align="center"><%IF oitem.FItemList(i).FSmallImage <> "" THEN%><img src="<%= oitem.FItemList(i).FSmallImage %>" width="50" height="50" border=0 alt=""><%END IF%></td>
		<td align="center"><% =oitem.FItemList(i).Fmakerid %></td>
	<td>
		<%If oitem.FItemList(i).Fitemdiv="21" Then%><font color="#0000ff">[��]</font><%End If%>&nbsp;<% =oitem.FItemList(i).Fitemname %>
		<b style="color:red"><%=chkIIF( tempSalePer = 0, "", "["&tempSalePer&"%]")%></b>
	</td>
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
		<img src="http://scm.10x10.co.kr/images/soldout_s.gif" width="30" height="12">
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
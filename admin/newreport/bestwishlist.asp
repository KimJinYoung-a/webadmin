<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ����Ʈ ���� ����Ʈ ���
' History : 2008.06.23 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/wishlist/bestwishlist_class.asp"-->
<!-- #include virtual="/lib/classes/displaycate/displaycateCls.asp"-->
<%
dim page ,cdl , sort , rectselly , ordertype
dim buy_date,buy_date1 , defaultnow, newitem, vCateCode
dim ipgocheck , mincash , maxcash, nvshop
	page = request("page")
		if page = "" then page = 1
	defaultnow = dateadd("d",-30,left(now(),10))		'���ó�¥���� -30��
	
	buy_date = request("buy_date")
	if buy_date = "" then		'�� ���� ���Ұ�� ������ �⺻��
		buy_date = left(defaultnow,4) &"-"&  mid(defaultnow,6,2) &"-"& mid(defaultnow,9,2)    	 
	end if
	
	buy_date1 = request("buy_date1")	
	if buy_date1 = "" then			'������ ���Ұ�� �������� �⺻��
		buy_date1 = left(now(),10)    	 
	end if	
	cdl = request("cdl")
	sort = request("sort")
		if sort = "" then sort = 100
	rectselly = request("rectselly")
	newitem = request("newitem") 
	vCateCode = Request("catecode")
	ordertype = request("ordertype")
		if ordertype = "" then ordertype = "select"		
	ipgocheck = request("ipgocheck")
	mincash = request("mincash")
	maxcash = request("maxcash")	
		if mincash = "" then mincash = 10000
		if maxcash = "" then maxcash = 20000
	nvshop = request("nvshop")
dim oip , i
set oip = new cwishlist
	oip.FPageSize = sort
	oip.FCurrPage = page
	oip.FRectSellY = rectselly
	oip.frectstartdate = buy_date
	oip.frectenddate = buy_date1
	oip.frectcdl = cdl
	oip.frectordertype = ordertype
	oip.frectipgocheck = ipgocheck
	oip.frectmincash = mincash
	oip.frectmaxcash = maxcash
	oip.frectnewitem = newitem
	oip.frectdisp1 = vCateCode
	oip.FRectNvshop	 = nvshop
	oip.fwishlist()
%>	
<script type="text/javascript" src="/js/jsCal/js/jscal2.js"></script>
<script type="text/javascript" src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script language="javascript">
function gopage(p){
	frm.page.value = p;
	frm.submit();
}

	function EnDisabledDateBox(comp){
		document.frm.mincash.disabled = !comp.checked;
		document.frm.maxcash.disabled = !comp.checked;
	}

	function reg(){
		if (document.frm.ipgocheck.checked){
			if (document.frm.mincash.value==''){
				alert('���ڸ� �Է����ּ���');
				document.frm.mincash.focus();
			}else if (document.frm.maxcash.value==''){
				alert('���ڸ� �Է����ּ���');
				document.frm.maxcash.focus();
			}else{
				frm.submit();
			}
			
		}else{
			frm.submit();
		}
	}

	function Check()
	{
	
	  if((event.keyCode<48) || (event.keyCode>57))
	    {return false;}
	}

</script>

<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="page" value="">
	<input type="hidden" name="menupos" value="<%=request("menupos")%>">
	<tr align="center" bgcolor="#FFFFFF">
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
		<td align="left">
			<input type="text" name="buy_date" value="<%=buy_date%>" class="formTxt" id="termSdt" style="width:100px" placeholder="������" />
			<input type="image" src="/images/admin_calendar.png" alt="�޷����� �˻�" id="ChkStart_trigger" onclick="return false;" />
			~
			<input type="text" name="buy_date1" value="<%=buy_date1%>" class="formTxt" id="termEdt" style="width:100px" placeholder="������" />
			<input type="image" src="/images/admin_calendar.png" alt="�޷����� �˻�" id="ChkEnd_trigger" onclick="return false;" />
			<script type="text/javascript">
				var CAL_Start = new Calendar({
					inputField : "termSdt", trigger    : "ChkStart_trigger",
					onSelect: function() {
						var date = Calendar.intToDate(this.selection.get());
						CAL_End.args.min = date;
						CAL_End.redraw();
						this.hide();
					}, bottomBar: true, dateFormat: "%Y-%m-%d"
				});
				var CAL_End = new Calendar({
					inputField : "termEdt", trigger    : "ChkEnd_trigger",
					onSelect: function() {
						var date = Calendar.intToDate(this.selection.get());
						CAL_Start.args.max = date;
						CAL_Start.redraw();
						this.hide();
					}, bottomBar: true, dateFormat: "%Y-%m-%d"
				});
			</script>

			����ī�װ� :  <% SelectBoxCategoryLarge cdl %>
			����ī�װ� : 
			<%
			Dim cDisp
			SET cDisp = New cDispCate
			cDisp.FCurrPage = 1
			cDisp.FPageSize = 2000
			cDisp.FRectDepth = 1
			cDisp.FRectUseYN = "Y"
			cDisp.GetDispCateList()
			
			If cDisp.FResultCount > 0 Then
				Response.Write "<select name=""catecode"" class=""select"" onChange=""frm.submit();"">" & vbCrLf
				Response.Write "<option value="""">����</option>" & vbCrLf
				For i=0 To cDisp.FResultCount-1
					Response.Write "<option value=""" & cDisp.FItemList(i).FCateCode & """ " & CHKIIF(CStr(vCateCode)=CStr(cDisp.FItemList(i).FCateCode),"selected","") & ">" & cDisp.FItemList(i).FCateName & "</option>"
				Next
				Response.Write "</select>&nbsp;&nbsp;&nbsp;"
			End If
			Set cDisp = Nothing
			%>
			���� ���̹������� : 
			<select name="nvshop" class="select">
				<option value="">-��ü-</option>
				<option value="nvshopY" <%= Chkiif(nvshop = "nvshopY", "selected", "") %> >����</option>
				<option value="nvshopN" <%= Chkiif(nvshop = "nvshopN", "selected", "") %> >������</option>
			</select>&nbsp;
			<input type="checkbox" name="rectselly" <% if rectselly="on" then response.write " checked" %>>�Ǹž����۸�
			<input type="checkbox" name="newitem" <% if newitem="on" then response.write " checked" %>>�Ż�ǰ��(2�ֳ���ϻ�ǰ)
		</td>	
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button" value="�˻�" onclick="reg();">
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF">
		<td align="left">			
			<input type="radio" name="ordertype" value="select" <% if ordertype="select" then response.write "checked" %>>�ֱټ���(��)��
			<input type="radio" name="ordertype" value="price" <% if ordertype="price" then response.write "checked" %>>���ݼ�
			&nbsp; �˻���: <select name="sort">
				<option value="10" <% if sort = "10" then response.write " selected" %>>10</option>
				<option value="20" <% if sort = "20" then response.write " selected" %>>20</option>
				<option value="30" <% if sort = "30" then response.write " selected" %>>30</option>
				<option value="40" <% if sort = "40" then response.write " selected" %>>40</option>
				<option value="50" <% if sort = "50" then response.write " selected" %>>50</option>
				<option value="60" <% if sort = "60" then response.write " selected" %>>60</option>
				<option value="70" <% if sort = "70" then response.write " selected" %>>70</option>
				<option value="80" <% if sort = "80" then response.write " selected" %>>80</option>
				<option value="90" <% if sort = "90" then response.write " selected" %>>90</option>																			
				<option value="100" <% if sort = "100" then response.write " selected" %>>100</option>
				<option value="100" <% if sort = "150" then response.write " selected" %>>150</option>
				<option value="100" <% if sort = "200" then response.write " selected" %>>200</option>					
			</select>
			<input type="checkbox" name="ipgocheck" value="on" <% if ipgocheck="on" then  response.write " checked" %> onclick="EnDisabledDateBox(this)">�������ǻ��
			<input type="text" name="mincash" size=10 value=<%=mincash%> onkeypress="return Check()">�̻�
			<input type="text" name="maxcash" size=10 value=<%=maxcash%> onkeypress="return Check()">�̸�
		</td>
	</tr>
	</form>
</table>
<!-- �˻� �� -->

<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<tr>
		<td align="left">
		</td>
		<td align="right">	
		</td>
	</tr>
</table>
<!-- �׼� �� -->

<!-- ����Ʈ ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15">
			<font color="red">* ���̹� ���� ������</font><br>
			- ���̹� ���ο� �����Ǿ� �ִ� ��ǰ �������� ���� 12�� 30�п� ������Ʈ�˴ϴ�.<br>
			- �˻����� �� �Ⱓ ������ ������� ���� �������θ� ǥ�ð� �˴ϴ�.
		</td>
	</tr>
	<% if oip.FresultCount>0 then %>
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15">
			�˻���� : <b><%= oip.FTotalCount %></b>
			&nbsp;
			������ : <b><%= page %>/ <%= oip.FTotalPage %></b>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
			<td align="center">����</td>
			<td align="center">Image</td>	
			<td align="center">Itemid</td>
			<td align="center">��ǰ��</td>
			<td align="center">brand</td>
			<td align="center">���̹�<br>����������</td>
			<td align="center">�ǸŰ�</td>
			<td align="center">���԰�</td>			
			<td align="center">�򰹼�</td>		
			<td align="center">��۱���</td>					
    </tr>
	<% for i = 0 to oip.FresultCount - 1 %>
	    <tr align="center" bgcolor="#FFFFFF">
			<td align="center"><%=(sort*(page-1)) + i+1 %></td>
			<td align="center">
				<img src="<%= oip.FItemList(i).fsmallimage %>" width=40 height=40 board=0>
			</td>	
			<td align="center"><a href="http://www.10x10.co.kr/shopping/category_prd.asp?itemid=<%= oip.FItemList(i).Fitemid %>" target="_blank" title="�̸�����"><%= oip.FItemList(i).fitemid %></a></td>
			<td align="center"><%= oip.FItemList(i).fitemname %></td>
			<td align="center"><%= oip.FItemList(i).fmakerid %></td>
			<td align="center">
			<%
				If oip.FItemList(i).FLowprice <> 0 Then
					response.write FormatNumber(oip.FItemList(i).FLowprice,0)
				End If
			%>
			</td>
			<td align="center"><%= FormatNumber(oip.FItemList(i).fsellcash,0) %></td>
			<td align="center"><%= FormatNumber(oip.FItemList(i).forgsuplycash,0) %></td>			
			<td align="center"><%= oip.FItemList(i).fitemid_count %></td>		
			<td align="center"><%= oip.FItemList(i).fmwdiv %></td>	      
	    </tr>   
	<% next %>
	
	<% else %>
	
		<tr bgcolor="#FFFFFF">
			<td colspan="15" align="center" class="page_link">[�˻������ �����ϴ�.]</td>
		</tr>
		
	<% end if %>
	
    <tr height="25" bgcolor="FFFFFF">
		<td colspan="15" align="center">
			<% if oip.HasPreScroll then %>
				<a href="javascript:gopage('<%= omd.StarScrollPage-1 %>');">[pre]</a>
			<% else %>
				[pre]
			<% end if %>
		
			<% for i=0 + oip.StartScrollPage to oip.FScrollCount + oip.StartScrollPage - 1 %>
				<% if i>oip.FTotalpage then Exit for %>
				<% if CStr(page)=CStr(i) then %>
				<font color="red">[<%= i %>]</font>
				<% else %>
				<a href="javascript:gopage('<%= i %>');">[<%= i %>]</a>
				<% end if %>
			<% next %>
		
			<% if oip.HasNextScroll then %>
				<a href="javascript:gopage('<%= i %>');">[next]</a>
			<% else %>
				[next]
			<% end if %>
		</td>
	</tr>
</table>
<script language='javascript'>
EnDisabledDateBox(document.frm.ipgocheck);
//UseIpCheck(document.frm.ipcheck);
</script>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->

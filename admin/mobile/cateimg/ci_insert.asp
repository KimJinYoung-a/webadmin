<%@ language=vbscript %>
<% option explicit %>
<%
'###############################################
' PageName : nb_insert.asp
' Discription : ����� ����Ʈ �˸����
' History : 2013.04.01 ����ȭ
'			2013.12.15 �ѿ�� ����
'###############################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/classes/mobile/catebanner.asp" -->
<%
Dim subImage1 , isusing , mode, oCateImgOne, idx, dispCate
Dim kword1 , kword2 , kword3 , kwordurl1 , kwordurl2 , kwordurl3
	idx = requestCheckvar(request("idx"),16)
	menupos = requestCheckvar(request("menupos"),10)

If idx = "" Then 
	mode = "add" 
Else 
	mode = "modify" 
End If 

set oCateImgOne = new CMainbanner
	oCateImgOne.FRectIdx = idx
	
	if idx<>"" then
		oCateImgOne.GetOneContents()
	end if
	
	if oCateImgOne.FResultCount > 0 then
		dispCate = oCateImgOne.FOneItem.Fcatecode
		isusing = oCateImgOne.FOneItem.Fisusing
		subImage1 = oCateImgOne.FOneItem.Fcateimg
		idx = oCateImgOne.FOneItem.fidx
		kword1 = oCateImgOne.FOneItem.fkword1
		kword2 = oCateImgOne.FOneItem.fkword2
		kword3 = oCateImgOne.FOneItem.fkword3
		kwordurl1 = oCateImgOne.FOneItem.fkwordurl1
		kwordurl2 = oCateImgOne.FOneItem.fkwordurl2
		kwordurl3 = oCateImgOne.FOneItem.fkwordurl3
	end if
set oCateImgOne = Nothing
%>
<link href="/js/jqueryui/css/jquery-ui.css" rel="stylesheet">
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/jqueryui/jquery-ui-1.10.2.custom.min.js"></script>
<script type='text/javascript'>

	function jsSubmit(){
		var frm = document.frm;
	
		if (!frm.disp.value){
			alert('ī�װ��� �������ּ���');
			frm.disp.focus();
			return;
		}

		if (confirm('���� �Ͻðڽ��ϱ�?')){
			//frm.target = "blank";
			frm.submit();
		}
	}
	
	function jsgolist(){
		self.location.href="/admin/mobile/cateimg/";
	}
	
	function putLinkText(key,gubun) {
		var frm = document.frm;
		var kword
		var urllink
		if (gubun == "1" )
		{
			urllink = frm.kwordurl1;
			kword = frm.kword1.value;
		}else if( gubun == "2"){
			urllink = frm.kwordurl2;
			kword = frm.kword2.value;
		}else{
			urllink = frm.kwordurl3;
			kword = frm.kword3.value;
		}
		switch(key) {
			case 'search':
				urllink.value='/search/search_result.asp?rect='+kword;
				break;
			case 'event':
				urllink.value='/event/eventmain.asp?eventid=�̺�Ʈ��ȣ';
				break;
			case 'itemid':
				urllink.value='/category/category_itemprd.asp?itemid=��ǰ�ڵ�';
				break;
			case 'category':
				urllink.value='/category/category_list.asp?disp=ī�װ�';
				break;
			case 'brand':
				urllink.value='/street/street_brand.asp?makerid=�귣����̵�';
				break;
		}
	}
</script>
<table width="100%" cellpadding="2" cellspacing="1" class="a" bgcolor="#3d3d3d">
<form name="frm" method="post" action="<%=uploadUrl%>/linkweb/mobile/doCateimage.asp" enctype="multipart/form-data" style="margin:0px;">
<input type="hidden" name="mode" value="<%=mode%>">
<input type="hidden" name="idx" value="<%=idx%>">
<input type="hidden" name="menupos" value="<%=menupos%>">

<% If mode = "modify" then%>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center" width="100">��ȣ</td>
	<td>
		<%= idx %>	<font color="red">�ؼ����ÿ��� ī�װ� ������ �Ұ��� �մϴ� . �̹����� ���� ���ּ��� ��</font>
	</td>
</tr>
<% End If %>

<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center" width="100">ī�װ�</td>
	<td>
		<!-- #include virtual="/common/module/dispCateSelectBox.asp"-->
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center">ī�װ��̹���</td>
	<td>
		<input type="file" name="subImage1" class="file" title="�̹��� #1" require="N" style="width:80%;" />
		<% if subImage1<>"" then %>
		<br>
		<img src="<%= subImage1 %>" width="100" /><br><%= subImage1 %>
		<% end if %>		
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center" width="10%">Ű����1</td>
	<td><input type="text" name="kword1" value="<%=kword1%>" size="20" maxlength="20"/></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center">Ű����1 URL</td>
	<td><input type="text" name="kwordurl1" size="80" value="<%=kwordurl1%>"/>
	<br/><br/>ex)
		<font color="#707070">
		- <span style="cursor:pointer" onClick="putLinkText('search','1')">�˻���� ��ũ : /search/search_result.asp?rect=<font color="darkred">�˻���</font></span><br>
		- <span style="cursor:pointer" onClick="putLinkText('event','1')">�̺�Ʈ ��ũ : /event/eventmain.asp?eventid=<font color="darkred">�̺�Ʈ�ڵ�</font></span><br>
		- <span style="cursor:pointer" onClick="putLinkText('itemid','1')">��ǰ�ڵ� ��ũ : /category/category_itemprd.asp?itemid=<font color="darkred">��ǰ�ڵ� (O)</font></span><br>
		- <span style="cursor:pointer" onClick="putLinkText('category','1')">ī�װ� ��ũ : /category/category_list.asp?disp=<font color="darkred">ī�װ�</font></span><br>
		- <span style="cursor:pointer" onClick="putLinkText('brand','1')">�귣����̵� ��ũ : /street/street_brand.asp?makerid=<font color="darkred">�귣����̵�</font></span>
		</font>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center" width="10%">Ű����2</td>
	<td><input type="text" name="kword2" value="<%=kword2%>" size="20" maxlength="20"/></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center">Ű����2 URL</td>
	<td><input type="text" name="kwordurl2" size="80" value="<%=kwordurl2%>"/>
	<br/><br/>ex)
		<font color="#707070">
		- <span style="cursor:pointer" onClick="putLinkText('search','2')">�˻���� ��ũ : /search/search_result.asp?rect=<font color="darkred">�˻���</font></span><br>
		- <span style="cursor:pointer" onClick="putLinkText('event','2')">�̺�Ʈ ��ũ : /event/eventmain.asp?eventid=<font color="darkred">�̺�Ʈ�ڵ�</font></span><br>
		- <span style="cursor:pointer" onClick="putLinkText('itemid','2')">��ǰ�ڵ� ��ũ : /category/category_itemprd.asp?itemid=<font color="darkred">��ǰ�ڵ� (O)</font></span><br>
		- <span style="cursor:pointer" onClick="putLinkText('category','2')">ī�װ� ��ũ : /category/category_list.asp?disp=<font color="darkred">ī�װ�</font></span><br>
		- <span style="cursor:pointer" onClick="putLinkText('brand','2')">�귣����̵� ��ũ : /street/street_brand.asp?makerid=<font color="darkred">�귣����̵�</font></span>
		</font>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center" width="10%">Ű����3</td>
	<td><input type="text" name="kword3" value="<%=kword3%>" size="20" maxlength="20"/></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center">Ű����3 URL</td>
	<td><input type="text" name="kwordurl3" size="80" value="<%=kwordurl3%>"/>
	<br/><br/>ex)
		<font color="#707070">
		- <span style="cursor:pointer" onClick="putLinkText('search','3')">�˻���� ��ũ : /search/search_result.asp?rect=<font color="darkred">�˻���</font></span><br>
		- <span style="cursor:pointer" onClick="putLinkText('event','3')">�̺�Ʈ ��ũ : /event/eventmain.asp?eventid=<font color="darkred">�̺�Ʈ�ڵ�</font></span><br>
		- <span style="cursor:pointer" onClick="putLinkText('itemid','3')">��ǰ�ڵ� ��ũ : /category/category_itemprd.asp?itemid=<font color="darkred">��ǰ�ڵ� (O)</font></span><br>
		- <span style="cursor:pointer" onClick="putLinkText('category','3')">ī�װ� ��ũ : /category/category_list.asp?disp=<font color="darkred">ī�װ�</font></span><br>
		- <span style="cursor:pointer" onClick="putLinkText('brand','3')">�귣����̵� ��ũ : /street/street_brand.asp?makerid=<font color="darkred">�귣����̵�</font></span>
		</font>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center">��뿩��</td>
	<td><div style="float:left;"><input type="radio" name="isusing" value="Y" <%=chkiif(isusing = "Y","checked","")%> checked />����� &nbsp;&nbsp;&nbsp; <input type="radio" name="isusing" value="N"  <%=chkiif(isusing = "N","checked","")%>/>������</div> <div style="float:right;margin-top:5px;margin-right:10px;"></div></td>
</tr>
<tr bgcolor="#FFFFFF" align="center">
    <td colspan="2"><input type="button" value=" �� �� " onClick="jsgolist();" class="button" /><input type="button" value=" �� �� " onClick="jsSubmit();" class="button" /></td>
</tr>
</form>
</table>

<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
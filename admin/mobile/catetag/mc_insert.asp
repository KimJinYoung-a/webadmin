<%@ language=vbscript %>
<% option explicit %>
<%
'###############################################
' PageName : mc_insert.asp
' Discription : ����� ����Ʈ ī�װ� �±�
' History : 2014-09-02 ����ȭ
'###############################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/classes/mobile/catetag.asp" -->
<%
Dim subImage1 , isusing , mode, oCatetagOne, idx, dispCate , appdiv , appcate
Dim kword1 , kword2 , kword3 , kwordurl1 , kwordurl2 , kwordurl3
	idx = requestCheckvar(request("idx"),16)
	menupos = requestCheckvar(request("menupos"),10)

If idx = "" Then 
	mode = "add" 
Else 
	mode = "modify" 
End If 

set oCatetagOne = new CMaincatetag
	oCatetagOne.FRectIdx = idx
	
	if idx<>"" then
		oCatetagOne.GetOneContents()
	end if
	
	if oCatetagOne.FResultCount > 0 then
		dispCate = oCatetagOne.FOneItem.Fcatecode
		isusing = oCatetagOne.FOneItem.Fisusing
		idx = oCatetagOne.FOneItem.fidx
		kword1 = oCatetagOne.FOneItem.fkword1
		kwordurl1 = oCatetagOne.FOneItem.fkwordurl1
		kwordurl2 = oCatetagOne.FOneItem.fkwordurl2
		appdiv = oCatetagOne.FOneItem.fappdiv
		appcate = oCatetagOne.FOneItem.fappcate
	end if
set oCatetagOne = Nothing
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
		self.location.href="/admin/mobile/catetag/?menupos=<%=menupos%>";
	}
	
	function putLinkText(key) {
		var frm = document.frm;
		var kword
		var urllink
			urllink = frm.kwordurl1;
			kword = frm.kword1.value;
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
				urllink.value='/category/category_list.asp?cdl=ī�װ�';
				break;
			case 'brand':
				urllink.value='/street/street_brand.asp?makerid=�귣����̵�';
				break;
		}
	}

	//url �ڵ� ����
	function chklink(v){
		if (v == "1"){
			document.frm.kwordurl2.value = "/apps/appCom/wish/web2014/category/category_itemprd.asp?itemid=��ǰ�ڵ�";
			$("#catesel").css("display","none");
			$("#kwordurl2").prop('disabled',false);
		}else if (v == "2"){
			document.frm.kwordurl2.value = "/apps/appcom/wish/web2014/event/eventmain.asp?eventid=�̺�Ʈ�ڵ�&rdsite=rdsite��(�ʼ��ƴ�)";
			$("#catesel").css("display","none");
			$("#kwordurl2").prop('disabled',false);
		}else if (v == "3"){
			document.frm.kwordurl2.value = "makerid=�귣���";
			$("#catesel").css("display","none");
			$("#kwordurl2").prop('disabled',false);
		}else if (v == "4"){
			chgDispCate2('');
			document.frm.kwordurl2.value = "cd1=&nm1=";
			$("#catesel").css("display","block");
			$("#kwordurl2").attr('readonly','readonly');
		}else{
			document.frm.kwordurl2.value = "APP URL ������ ���� ���ּ���.";
			$("#catesel").css("display","none");
			$("#kwordurl2").prop('disabled',false);
		}
	}
</script>
<script>
function chgDispCate2(dc) {
	$.ajax({
		url: "dispCateSelectBox_response.asp?disp="+dc,
		cache: false,
		async: false,
		success: function(message) {
			// ���� �ֱ�
			$("#lyrDispCtBox2").empty().html(message);
			if (dc.length == 3){
				document.frm.kwordurl2.value = $("#dispcateval1 option:selected").val()+"||"+$("#dispcateval1 option:selected").text();
				$("#appcate").val(dc);
			}else if (dc.length == 6){
				document.frm.kwordurl2.value = $("#dispcateval1 option:selected").val()+"||"+$("#dispcateval2 option:selected").val()+"||"+$("#dispcateval1 option:selected").text()+"||"+$("#dispcateval2 option:selected").text();
				$("#appcate").val(dc);
			}else if (dc.length == 9){
				document.frm.kwordurl2.value = $("#dispcateval1 option:selected").val()+"||"+$("#dispcateval2 option:selected").val()+"||"+$("#dispcateval3 option:selected").val()+"||"+$("#dispcateval1 option:selected").text()+"||"+$("#dispcateval2 option:selected").text()+"||"+$("#dispcateval3 option:selected").text();
				$("#appcate").val(dc);
			}else{
				
			}

		}
	});
}
$(function(){
	chgDispCate2('<%=appcate%>');
});
</script>
<table width="100%" cellpadding="2" cellspacing="1" class="a" bgcolor="#3d3d3d">
<form name="frm" method="post" action="docatetag.asp" onSubmit="return jsRegCode();">
<input type="hidden" name="mode" value="<%=mode%>"/>
<input type="hidden" name="idx" value="<%=idx%>"/>
<input type="hidden" name="menupos" value="<%=menupos%>"/>
<input type="hidden" name="appcate" id="appcate"/>
<% If mode = "modify" then%>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center" width="100">��ȣ</td>
	<td>
		<%= idx %>
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
	<td bgcolor="#FFF999" align="center" width="10%">�α� Ű����</td>
	<td><input type="text" name="kword1" value="<%=kword1%>" size="20" maxlength="20"/></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center">��������� URL</td>
	<td><input type="text" name="kwordurl1" size="80" value="<%=kwordurl1%>"/>
	<br/><br/>ex)
		<font color="#707070">
		- <span style="cursor:pointer" onClick="putLinkText('search')">�˻���� ��ũ : /search/search_result.asp?rect=<font color="darkred">�˻���</font></span><br>
		- <span style="cursor:pointer" onClick="putLinkText('event')">�̺�Ʈ ��ũ : /event/eventmain.asp?eventid=<font color="darkred">�̺�Ʈ�ڵ�</font></span><br>
		- <span style="cursor:pointer" onClick="putLinkText('itemid')">��ǰ�ڵ� ��ũ : /category/category_itemprd.asp?itemid=<font color="darkred">��ǰ�ڵ� (O)</font></span><br>
		- <span style="cursor:pointer" onClick="putLinkText('category')">ī�װ� ��ũ : /category/category_list.asp?cdl=<font color="darkred">ī�װ�</font></span><br>
		- <span style="cursor:pointer" onClick="putLinkText('brand')">�귣����̵� ��ũ : /street/street_brand.asp?makerid=<font color="darkred">�귣����̵�</font></span>
		</font>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center">APP�� URL</td>
	<td>
		<table width="100%" cellpadding="3" cellspacing="1" class="a" bgcolor="#3d3d3d">
			<tr>
				<td bgcolor="#FFF999" width="100" align="center">APP URL ����</td>
				<td bgcolor="#FFFFFF">
					<select name='appdiv' class='select' onchange="chklink(this.value);">
						<option value="0">�����ϼ���</option>
						<option value="1" <% if appdiv = "1" then response.write " selected" %>>��ǰ��</option>
						<option value="2" <% if appdiv = "2" then response.write " selected" %>>�̺�Ʈ</option>
						<option value="3" <% if appdiv = "3" then response.write " selected" %>>�귣��</option>
						<option value="4" <% if appdiv = "4" then response.write " selected" %>>ī�װ�</option>
					</select>
				</td>
			</tr>
			<tr id="catesel" style="display:<%=chkiif(idx<>"" And appdiv = "4","block","none")%>">
				<td bgcolor="#FFF999" width="100" align="center">����ī�װ� ����</td>
				<td bgcolor="#FFFFFF">
					<span id="lyrDispCtBox2"></span>
				</td>
			</tr>
			<tr>
				<td bgcolor="#FFF999" width="100" align="center">�ڵ峻��</td>
				<td bgcolor="#FFFFFF"><textarea name="kwordurl2" class="textarea" id="kwordurl2" style="width:100%; height:40px;"><%=kwordurl2%></textarea></td>
			</tr>
		</table>
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
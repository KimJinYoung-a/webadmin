<%@ language=vbscript %>
<% option explicit %>
<%
'###############################################
' PageName : tkw_insert.asp
' Discription : ����� GNB top keyword
' History : 2015-09-16 ����ȭ
'###############################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/classes/mobile/topkeyword.asp" -->
<!-- #include virtual="/lib/classes/mobile/topcatecodeCls.asp" -->
<%
Dim idx , subImage1 , isusing , mode , gcode
Dim srcSDT , srcEDT 
Dim mainStartDate, mainEndDate
Dim sDt, sTm, eDt, eTm
Dim sortnum  , prevDate , ordertext
Dim ktitle , kword , kcontents , kurl_mo , kurl_app , appdiv , appcate
Dim itemid , itemName , smallImage
	idx = requestCheckvar(request("idx"),16)
	gcode = requestCheckvar(request("gcode"),3)
	srcSDT = request("sDt")
	srcEDT = request("eDt")
	prevDate = request("prevDate")

If idx = "" Then 
	mode = "add" 
Else 
	mode = "modify" 
End If 

If idx <> "" then
	dim tkwobannerOne
	set tkwobannerOne = new CMainbanner
	tkwobannerOne.FRectIdx = idx
	tkwobannerOne.GetOneContents()

	kword				=	tkwobannerOne.FOneItem.Fkword
	ktitle				=	tkwobannerOne.FOneItem.Fktitle
	kcontents			=	tkwobannerOne.FOneItem.Fkcontents
	kurl_mo				=	tkwobannerOne.FOneItem.Fkurl_mo
	kurl_app			=	tkwobannerOne.FOneItem.Fkurl_app
	appdiv				=	tkwobannerOne.FOneItem.Fappdiv
	appcate				=	tkwobannerOne.FOneItem.fappcate
	sortnum				=	tkwobannerOne.FOneItem.Fsortnum
	mainStartDate		=	tkwobannerOne.FOneItem.Fstartdate
	mainEndDate			=	tkwobannerOne.FOneItem.Fenddate 
	isusing				=	tkwobannerOne.FOneItem.Fisusing
	subImage1			=	tkwobannerOne.FOneItem.Fkwimg
	ordertext			=	tkwobannerOne.FOneItem.Fordertext
	itemid				=	tkwobannerOne.FOneItem.Fitemid
	itemname			=	tkwobannerOne.FOneItem.Fitemname
	smallimage			=	tkwobannerOne.FOneItem.Fsmallimage
	gcode				=	tkwobannerOne.FOneItem.Fgnbcode

	set tkwobannerOne = Nothing
End If 

if Not(mainStartDate="" or isNull(mainStartDate)) then
	sDt = left(mainStartDate,10)
	sTm = Num2Str(hour(mainStartDate),2,"0","R") &":"& Num2Str(minute(mainStartDate),2,"0","R") &":"& Num2Str(second(mainStartDate),2,"0","R")
else
	if srcSDT<>"" then
		sDt = left(srcSDT,10)
	else
		sDt = date
	end if
	sTm = "00:00:00"
end if

if Not(mainEndDate="" or isNull(mainEndDate)) then
	eDt = left(mainEndDate,10)
	eTm = Num2Str(hour(mainEndDate),2,"0","R") &":"& Num2Str(minute(mainEndDate),2,"0","R") &":"& Num2Str(second(mainEndDate),2,"0","R")
else
	if srcEDT<>"" then
		eDt = left(srcEDT,10)
	else
		eDt = date
	end if
	eTm = "23:59:59"
end if
%>
<link href="/js/jqueryui/css/jquery-ui.css" rel="stylesheet">
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/jqueryui/jquery-ui-1.10.2.custom.min.js"></script>
<script type='text/javascript'>
	function jsSubmit(){
		var frm = document.frm;

		if (!frm.gcode.value)
		{
			alert('���� GNB ������ ���� ���ּ���.');
			frm.gcode.focus();
			return;
		}

//		if (!frm.kword.value)
//		{
//			alert('Ű���带 �Է����ּ���.');
//			frm.kword.focus();
//			return;
//		}

		if (!frm.ktitle.value)
		{
			alert('������ �Է����ּ���.');
			frm.ktitle.focus();
			return;
		}
//
//		if (!frm.kcontents.value)
//		{
//			alert('������ �Է����ּ���.');
//			frm.kcontents.focus();
//			return;
//		}

		if (!frm.kurl_mo.value)
		{
			alert('����� URL�� �Է����ּ���.');
			frm.kurl_mo.focus();
			return;
		}

		if (!frm.kurl_app.value)
		{
			alert('�� URL�� �Է����ּ���.');
			frm.kurl_app.focus();
			return;
		}
	
		if (confirm('���� �Ͻðڽ��ϱ�?')){
			//frm.target = "blank";
			frm.submit();
		}
	}
	function jsgolist(){
		self.location.href="/admin/mobile/topkeyword/";
	}
	$(function(){
	//�޷´�ȭâ ����
	var arrDayMin = ["��","��","ȭ","��","��","��","��"];
	var arrMonth = ["1��","2��","3��","4��","5��","6��","7��","8��","9��","10��","11��","12��"];
    $("#sDt").datepicker({
		dateFormat: "yy-mm-dd",
		prevText: '������', nextText: '������', yearSuffix: '��',
		dayNamesMin: arrDayMin,
		monthNames: arrMonth,
		showMonthAfterYear: true,
    	numberOfMonths: 1,
    	showCurrentAtPos: 0,
      	showOn: "button",
      	<% if Idx<>"" then %>maxDate: "<%=eDt%>",<% end if %>
    	onClose: function( selectedDate ) {
    		$( "#eDt" ).datepicker( "option", "minDate", selectedDate );
    	}
    });
    $("#eDt").datepicker({
		dateFormat: "yy-mm-dd",
		prevText: '������', nextText: '������', yearSuffix: '��',
		dayNamesMin: arrDayMin,
		monthNames: arrMonth,
		showMonthAfterYear: true,
    	numberOfMonths: 2,
      	showOn: "button",
      	<% if Idx<>"" then %>minDate: "<%=sDt%>",<% end if %>
    	onClose: function( selectedDate ) {
    		$( "#sDt" ).datepicker( "option", "maxDate", selectedDate );
    	}
    });
});

function putLinkText(key,gubun) {
	var frm = document.frm;
	var urllink
	if (gubun == "1" )
	{
		urllink = frm.kurl_mo;
	}

	switch(key) {
//		case 'search':
//			urllink.value='/search/search_result.asp?rect=�˻���';
//			break;
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

//url �ڵ� ����
function chklink(v){
	if (v == "1"){
		document.frm.kurl_app.value = "/apps/appCom/wish/web2014/category/category_itemprd.asp?itemid=��ǰ�ڵ�";
		alert('Mobile URL ���� ����!');
		$("#catesel").css("display","none");
		$("#kurl_app").prop('disabled',false);
	}else if (v == "2"){
		document.frm.kurl_app.value = "/apps/appcom/wish/web2014/event/eventmain.asp?eventid=�̺�Ʈ�ڵ�&rdsite=rdsite��(�ʼ��ƴ�)";
		alert('Mobile URL ���� ����!');
		$("#catesel").css("display","none");
		$("#kurl_app").prop('disabled',false);
	}else if (v == "3"){
		document.frm.kurl_app.value = "makerid=�귣���";
		alert('Mobile URL ���� ����!');
		$("#catesel").css("display","none");
		$("#kurl_app").prop('disabled',false);
	}else if (v == "4"){
		chgDispCate2('');
		document.frm.kurl_app.value = "cd1=&nm1=";
		$("#catesel").css("display","block");
		$("#kurl_app").attr('readonly','readonly');
	}else{
		document.frm.kurl_app.value = "APP URL ������ ���� ���ּ���.";
		$("#catesel").css("display","none");
		$("#kurl_app").prop('disabled',false);
	}
}

function chgDispCate2(dc) {
	$.ajax({
		url: "/admin/mobile/catetag/dispCateSelectBox_response.asp?disp="+dc,
		cache: false,
		async: false,
		success: function(message) {
			// ���� �ֱ�
			$("#lyrDispCtBox2").empty().html(message);
			if (dc.length == 3){
				document.frm.kurl_app.value = $("#dispcateval1 option:selected").val()+"||"+$("#dispcateval1 option:selected").text();
				$("#appcate").val(dc);
			}else if (dc.length == 6){
				document.frm.kurl_app.value = $("#dispcateval1 option:selected").val()+"||"+$("#dispcateval2 option:selected").val()+"||"+$("#dispcateval1 option:selected").text()+"||"+$("#dispcateval2 option:selected").text();
				$("#appcate").val(dc);
			}else if (dc.length == 9){
				document.frm.kurl_app.value = $("#dispcateval1 option:selected").val()+"||"+$("#dispcateval2 option:selected").val()+"||"+$("#dispcateval3 option:selected").val()+"||"+$("#dispcateval1 option:selected").text()+"||"+$("#dispcateval2 option:selected").text()+"||"+$("#dispcateval3 option:selected").text();
				$("#appcate").val(dc);
			}else{
				
			}

		}
	});
}
$(function(){
	<% if appdiv ="4" then %>
	chgDispCate2('<%=appcate%>');
	<% end if %>
});

// ��ǰ���� ����
function fnGetItemInfo(iid) {
	$.ajax({
		type: "GET",
		url: "/admin/sitemaster/wcms/act_iteminfo.asp?itemid="+iid,
		dataType: "xml",
		cache: false,
		async: false,
		timeout: 5000,
		beforeSend: function(x) {
			if(x && x.overrideMimeType) {
				x.overrideMimeType("text/xml;charset=euc-kr");
			}
		},
		success: function(xml) {
			if($(xml).find("itemInfo").find("item").length>0) {
				var rst = "<img src='" + $(xml).find("itemInfo").find("item").find("smallImage").text() + "' height='50' /><br/>"
					rst += $(xml).find("itemInfo").find("item").find("itemname").text();
				$("#lyItemInfo").fadeIn();
				$("#lyItemInfo").html(rst);
			} else {
				$("#lyItemInfo").fadeOut();
			}
		},
		error: function(xhr, status, error) {
			$("#lyItemInfo").fadeOut();
			/*alert(xhr + '\n' + status + '\n' + error);*/
		}
	});
}

</script>
<table width="900" cellpadding="2" cellspacing="1" class="a" bgcolor="#3d3d3d">
<form name="frm" method="post" action="<%=uploadUrl%>/linkweb/mobile/tkwbanner_proc.asp" enctype="multipart/form-data" style="margin:0px;">
<input type="hidden" name="mode" value="<%=mode%>">
<input type="hidden" name="adminid" value="<%=session("ssBctId")%>">
<input type="hidden" name="idx" value="<%=idx%>">
<input type="hidden" name="prevDate" value="<%=prevDate%>">
<input type="hidden" name="appcate" id="appcate"/>
<tr bgcolor="#FFFFFF">
    <td bgcolor="#FFF999" align="center" width="20">����Ⱓ</td>
    <td >
		<input type="text" id="sDt" name="StartDate" size="10" value="<%=sDt%>" />
		<input type="text" name="sTm" size="8" value="<%=sTm%>" /> ~
		<input type="text" id="eDt" name="EndDate" size="10" value="<%=eDt%>" />
		<input type="text" name="eTm" size="8" value="<%=eTm%>" />
    </td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center">���� GNB����</td>
	<td ><% Call drawSelectBoxGNB("gcode" , gcode) %></td>
</tr>
<!-- <tr bgcolor="#FFFFFF"> -->
<!-- 	<td bgcolor="#FFF999" align="center">Ű����</td> -->
<!-- 	<td ><input type="text" name="kword" size="50" value="<%=kword%>"/></td> -->
<!-- </tr> -->
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center">����</td>
	<td ><input type="text" name="ktitle" size="50" value="<%=ktitle%>"/></td>
</tr>
<!-- <tr bgcolor="#FFFFFF"> -->
<!-- 	<td bgcolor="#FFF999" align="center">����</td> -->
<!-- 	<td ><textarea name="kcontents" cols="50" rows="4"/><%=kcontents%></textarea></td> -->
<!-- </tr> -->
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center" colspan="2">��ǰ�ڵ�� �̹����� <span style="color:red">�Ѱ��� �̻�</span> �Է� - �̹����� �켱���� �ѷ����ϴ�.</td>
</tr>
<tr bgcolor="#FFFFFF">
    <td bgcolor="#FFF999" align="center">��ǰ�ڵ�</td>
    <td colspan="3">
        <input type="text" name="itemid" value="<%= itemid %>" size="8" maxlength="8" class="text" require="N" onblur="fnGetItemInfo(this.value)" title="��ǰ�ڵ�" />
        <div id="lyItemInfo" style="display:<%=chkIIF(itemid="","none","")%>;">
        <%
        	if Not(itemName="" or isNull(itemName)) then
        		Response.Write "<img src='" & smallImage & "' height='50' /><br/>"
        		Response.Write itemName
        	end if
        %>
        </div>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center" >�̹���</td>
	<td>
		<input type="file" name="subImage1" class="file" title="�̹��� #1" require="N" style="width:80%;" />
		<% if subImage1<>"" then %>
		<br>
		<img src="<%= subImage1 %>" width="100" /><br><%= subImage1 %>
		[<span style="color:red">�̹�������</span>] --&gt; <input type="checkbox" name="delimg" value="1"/>
		<% end if %>		
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center">����� URL</td>
	<td ><input type="text" name="kurl_mo" size="80" value="<%=kurl_mo%>"/>
	<br/><br/>ex)
		<font color="#707070">
		<!-- - <span style="cursor:pointer" onClick="putLinkText('search','1')">�˻���� ��ũ : /search/search_result.asp?rect=<font color="darkred">�˻���</font></span><br> -->
		- <span style="cursor:pointer" onClick="putLinkText('event','1')">�̺�Ʈ ��ũ : /event/eventmain.asp?eventid=<font color="darkred">�̺�Ʈ�ڵ�</font></span><br>
		- <span style="cursor:pointer" onClick="putLinkText('itemid','1')">��ǰ�ڵ� ��ũ : /category/category_itemprd.asp?itemid=<font color="darkred">��ǰ�ڵ� (O)</font></span><br>
		- <span style="cursor:pointer" onClick="putLinkText('category','1')">ī�װ� ��ũ : /category/category_list.asp?disp=<font color="darkred">ī�װ�</font></span><br>
		- <span style="cursor:pointer" onClick="putLinkText('brand','1')">�귣����̵� ��ũ : /street/street_brand.asp?makerid=<font color="darkred">�귣����̵�</font></span>
		</font>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center">�� URL</td>
	<td >
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
				<td bgcolor="#FFFFFF"><textarea name="kurl_app" class="textarea" id="kurl_app" style="width:100%; height:40px;"><%=kurl_app%></textarea></td>
			</tr>
		</table>
	</td>
</tr>

<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center">��뿩��</td>
	<td ><div style="float:left;"><input type="radio" name="isusing" value="Y" <%=chkiif(isusing = "Y","checked","")%> checked />����� &nbsp;&nbsp;&nbsp; <input type="radio" name="isusing" value="N"  <%=chkiif(isusing = "N","checked","")%>/>������</div> <div style="float:right;margin-top:5px;margin-right:10px;"></div></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center">���Ĺ�ȣ</td>
	<td ><input type="text" name="sortnum" value="<%=chkiif(sortnum="","0",sortnum)%>" size="2"/></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center">�۾��� ���û���</td>
	<td ><textarea name="ordertext" cols="80" rows="8"/><%=ordertext%></textarea></td>
</tr>
<tr bgcolor="#FFFFFF" align="center">
    <td colspan="2"><input type="button" value=" �� �� " onClick="jsgolist();"/><input type="button" value=" �� �� " onClick="jsSubmit();"/></td>
</tr>
</form>
</table>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
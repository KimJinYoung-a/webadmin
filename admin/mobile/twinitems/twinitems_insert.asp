<%@ language=vbscript %>
<% option explicit %>
<%
'###############################################
' PageName : itemtwins_insert.asp
' Discription : ����� ��ǰ���
' History : 2017-08-02 ����ȭ ����
'###############################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/classes/mobile/today_twinitemsCls.asp" -->
<%
Dim mode
Dim idx
Dim mainStartDate, mainEndDate
Dim sDt, sTm, eDt, eTm
Dim ordertext , isusing
Dim stdt , eddt
Dim L_img , L_maincopy , L_itemname	, L_itemid , L_newbest , R_img , R_maincopy	, R_itemname , R_itemid	, R_newbest , iteminfo

idx = requestCheckvar(request("idx"),16)

If idx = "" Then 
	mode = "add" 
Else 
	mode = "modify" 
End If 


'// ������
If idx <> "" then
	dim twinitemsOne
	set twinitemsOne = new CMainbanner
	twinitemsOne.FRectIdx = idx
	twinitemsOne.GetOneContents()

	mainStartDate		=	twinitemsOne.FOneItem.Fstartdate
	mainEndDate			=	twinitemsOne.FOneItem.Fenddate 
	isusing				=	twinitemsOne.FOneItem.Fisusing
	ordertext			=	twinitemsOne.FOneItem.Fordertext
	L_img				=	twinitemsOne.FOneItem.FL_img		
	L_maincopy			=	twinitemsOne.FOneItem.FL_maincopy
	L_itemname			=	twinitemsOne.FOneItem.FL_itemname
	L_itemid			=	twinitemsOne.FOneItem.FL_itemid	
	L_newbest			=	twinitemsOne.FOneItem.FL_newbest	
	R_img				=	twinitemsOne.FOneItem.FR_img		
	R_maincopy			=	twinitemsOne.FOneItem.FR_maincopy
	R_itemname			=	twinitemsOne.FOneItem.FR_itemname
	R_itemid			=	twinitemsOne.FOneItem.FR_itemid	
	R_newbest			=	twinitemsOne.FOneItem.FR_newbest	
	iteminfo			=	twinitemsOne.FOneItem.Fiteminfo	
	set twinitemsOne = Nothing

	Dim ii
	if not isnull(iteminfo) then 
		If ubound(Split(iteminfo,"^^")) > 0 Then ' �̹��� 2�� ����
			For ii = 0 To ubound(Split(iteminfo,"^^"))
				If CStr(L_itemid) = CStr(Split(Split(iteminfo,",")(ii),"|")(0)) And L_img = (staticImgUrl & "/mobile/twinitems") Then
					L_img =  webImgUrl & "/image/icon1/" & GetImageSubFolderByItemid(L_itemid) & "/" & Split(Split(iteminfo,",")(ii),"|")(2)
				End If

				If CStr(R_itemid) = CStr(Split(Split(iteminfo,",")(ii),"|")(0)) And R_img = (staticImgUrl & "/mobile/twinitems") Then
					R_img =  webImgUrl & "/image/icon1/" & GetImageSubFolderByItemid(R_itemid) & "/" & Split(Split(iteminfo,",")(ii),"|")(2)
				End If
			Next 
		End If 
	end if
End If 

dim dateOption
dateOption = request("dateoption")

if Not(mainStartDate="" or isNull(mainStartDate)) then
	sDt = left(mainStartDate,10)
	sTm = Num2Str(hour(mainStartDate),2,"0","R") &":"& Num2Str(minute(mainStartDate),2,"0","R") &":"& Num2Str(second(mainStartDate),2,"0","R")
elseif dateOption <> "" then
	sDt = dateOption
	sTm = "00:00:00"
else
	sDt = date
	sTm = "00:00:00"
end if

if Not(mainEndDate="" or isNull(mainEndDate)) then
	eDt = left(mainEndDate,10)
	eTm = Num2Str(hour(mainEndDate),2,"0","R") &":"& Num2Str(minute(mainEndDate),2,"0","R") &":"& Num2Str(second(mainEndDate),2,"0","R")
elseif dateOption <> "" then	
	eDt = dateOption
	eTm = "23:59:59"
else
	eDt = date
	eTm = "23:59:59"
end If

%>
<link href="/js/jqueryui/css/jquery-ui.css" rel="stylesheet">
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/jqueryui/jquery-ui-1.10.2.custom.min.js"></script>
<script type='text/javascript'>
	function jsSubmit(){
		var frm = document.frm;

		if (frm.L_itemid.value == ""){
			alert("���� ��ǰ�ڵ带 �־��ּ���.");
			return;
		}

		if (frm.L_itemname.value == ""){
			alert("���� ��ǰ�̸��� �־��ּ���.");
			frm.L_itemname.focus();
			return;
		}

		if (frm.L_maincopy.value == ""){
			alert("���� ����ī�Ǹ� �־��ּ���.");
			frm.L_maincopy.focus();
			return;
		}

		if (frm.R_itemid.value == ""){
			alert("���� ��ǰ�ڵ带 �־��ּ���.");
			return;
		}

		if (frm.R_itemname.value == ""){
			alert("���� ��ǰ�̸��� �־��ּ���.");
			frm.R_itemname.focus();
			return;
		}

		if (frm.R_maincopy.value == ""){
			alert("���� ����ī�Ǹ� �־��ּ���.");
			frm.R_maincopy.focus();
			return;
		}

		if (confirm('���� �Ͻðڽ��ϱ�?')){
			//frm.target = "blank";
			frm.submit();
		}
	}
	function jsgolist(){
		self.location.href="/admin/mobile/twinitems/";
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
    	numberOfMonths: 2,
    	showCurrentAtPos: 1,
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


function chgtype(v){
	if (v == "1"){
		$("#additem1").css("display","none");
		$("#additem2").css("display","none");
		$("#additem3").css("display","none");
	}else{
		$("#additem1").css("display","");
		$("#additem2").css("display","");
	}
}

// ��ǰ���� ����
function fnGetItemInfo(iid,v) {
	if (iid != "")
	{
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
					var rst = $(xml).find("itemInfo").find("item").find("itemname").text();
					//$("#lyItemInfo"+v).fadeIn();
					$("#lyItemInfo"+v).text(rst);
				} else {
					//$("#lyItemInfo"+v).fadeOut();
				}
			},
			error: function(xhr, status, error) {
				alert("��ǰ��ȣ�� �ٽ� �־� �ּ���");
				return; 
				// $("#lyItemInfo"+v).fadeOut();
				/*alert(xhr + '\n' + status + '\n' + error);*/
			}
		});
	}
}
</script>
<table width="90%" cellpadding="2" cellspacing="1" class="a" bgcolor="#3d3d3d">
<form name="frm" method="post" action="<%=uploadUrl%>/linkweb/mobile/twinitems_proc.asp" enctype="multipart/form-data" style="margin:0px;">
<input type="hidden" name="mode" value="<%=mode%>">
<input type="hidden" name="adminid" value="<%=session("ssBctId")%>">
<input type="hidden" name="idx" value="<%=idx%>">
<input type="hidden" name="menupos" value="<%=menupos%>">
<tr bgcolor="#FFFFFF">
	<td colspan="4" bgcolor="#FFF999" align="center"><%=chkiif(mode="add","�Է������� �Դϴ�.","���������� �Դϴ�.")%></td>
</tr>
<tr bgcolor="#FFFFFF">
    <td bgcolor="#FFF999" align="center" width="5%">����Ⱓ</td>
    <td colspan="3">
		<input type="text" id="sDt" name="StartDate" size="10" value="<%=sDt%>" />
		<input type="text" name="sTm" size="8" value="<%=sTm%>" /> ~
		<input type="text" id="eDt" name="EndDate" size="10" value="<%=eDt%>" />
		<input type="text" name="eTm" size="8" value="<%=eTm%>" />
    </td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center" width="5%">������<br/><br/>�������</td>
	<td width="40%">
		<table width="100%" cellpadding="0" cellspacing="0" class="a">
			<tr>
				<td align="left">
					��ǰ�ڵ� : <input type="text" name="L_itemid" value="<%=L_itemid%>" size="8" maxlength="8" class="text" require="N" onblur="fnGetItemInfo(this.value,'1')" title="��ǰ�ڵ�" />
					<br/>
					<% If L_img <> "" Then %>
					<img src="<%=L_img%>" width="120" height="120"/>
					<% Else %>
					<img src="/images/admin_login_logo2.png" width="120" height="120" /></br>�̹����� ��� ���ּ���.
					<% End If %>
				</td>
				<td align="right">
					����ī�� : <input type="text" name="L_maincopy" value="<%=L_maincopy%>" maxlength="10" size="20"/>
					<br><font color="red"><strong>�� �ִ� 10�� ���� ��</strong></font>
					<br/><br/>
					��ǰ��&nbsp; : &nbsp;&nbsp;<input type="text" name="L_itemname" value="<%=L_itemname%>" size="20" maxlength="8"/>
					<br><font color="red"><strong>�� �ִ� 8�� ���� ��</strong></font>
					<br/><br/><span id="lyItemInfo1"></span>
				</td>
			</tr>
			<tr>
				<td colspan="2">
					<%=L_img%><br/>
					�̹��� ��� : <input type="file" name="L_img" class="file" title="�̺�Ʈ #1" require="N" style="width:80%;" />
				</td>
			</tr>
			<tr>
				<td>
					<input type="radio" name="L_newbest" value="0" checked/> ������&nbsp;&nbsp;&nbsp;<input type="radio" name="L_newbest" value="1" <%=chkiif(L_newbest="1","checked","")%>/> NEW&nbsp;&nbsp;&nbsp; <input type="radio" name="L_newbest" value="2" <%=chkiif(L_newbest="2","checked","")%>/> BEST
				</td>
				<td align="right" width="50%" style="padding-right:30px;">
					<input type="checkbox" name="L_delimg" value="Y" id="L_delimg"/> <label for="L_delimg">�̹��� ����</label>
				</td>
			</tr>
		</table>
	</td>
	<td bgcolor="#FFF999" align="center" width="5%">������<br/><br/>�������</td>
	<td width="40%">
		<table width="100%" cellpadding="0" cellspacing="0" class="a">
			<tr>
				<td align="left">
					����ī�� : <input type="text" name="R_maincopy" value="<%=R_maincopy%>"  maxlength="10" size="20"/>
					<br><font color="red"><strong>�� �ִ� 10�� ���� ��</strong></font>
					<br/><br/>
					��ǰ��&nbsp; : <input type="text" name="R_itemname" value="<%=R_itemname%>" size="20" maxlength="8" />
					<br><font color="red"><strong>�� �ִ� 8�� ���� ��</strong></font>
					<br/><br/><span id="lyItemInfo2"></span>
				</td>
				<td align="right">
					��ǰ�ڵ� : <input type="text" name="R_itemid" value="<%=R_itemid%>" size="8" maxlength="8" class="text" require="N" onblur="fnGetItemInfo(this.value,'2')" title="��ǰ�ڵ�" />
					<br/>
					<% If R_img <> "" Then %>
					<img src="<%=R_img%>" width="120" height="120"/>
					<% Else %>
					<img src="/images/admin_login_logo2.png" width="120" height="120" /></br>�̹����� ��� ���ּ���.
					<% End If %>
				</td>
			</tr>
			<tr>
				<td colspan="2">
					<%=R_img%><br/>
					�̹��� ��� : <input type="file" name="R_img" class="file" title="�̺�Ʈ #1" require="N" style="width:80%;" />
				</td>
			</tr>
			<tr>
				<td>
					<input type="radio" name="R_newbest" value="0" checked/> ������&nbsp;&nbsp;&nbsp;<input type="radio" name="R_newbest" value="1" <%=chkiif(R_newbest="1","checked","")%>/> NEW&nbsp;&nbsp;&nbsp; <input type="radio" name="R_newbest" value="2" <%=chkiif(R_newbest="2","checked","")%>/> BEST
				</td>
				<td align="right" style="padding-right:30px;">
					<input type="checkbox" name="R_delimg" value="Y" id="R_delimg"/> <label for="R_delimg">�̹��� ����</label>
				</td>
			</tr>
		</table>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center">��뿩��</td>
	<td colspan="3"><div style="float:left;"><input type="radio" name="isusing" value="Y" <%=chkiif(isusing = "Y","checked","")%> checked />����� &nbsp;&nbsp;&nbsp; <input type="radio" name="isusing" value="N"  <%=chkiif(isusing = "N","checked","")%>/>������</div> <div style="float:right;margin-top:5px;margin-right:10px;"></div></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center">�۾��� ���û���</td>
	<td colspan="3"><textarea name="ordertext" cols="80" rows="8"/><%=ordertext%></textarea></td>
</tr>
<tr bgcolor="#FFFFFF" align="center">
    <td colspan="4"><input type="button" value=" �� �� " onClick="jsgolist();"/><input type="button" value=" �� �� " onClick="jsSubmit();"/></td>
</tr>
</form>
</table>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
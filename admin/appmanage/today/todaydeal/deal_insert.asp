<%@ language=vbscript %>
<% option explicit %>
<%
'###############################################
' PageName : deal_insert.asp
' Discription : ����� dealbanner_new
' History : 2014.06.23 ����ȭ
'###############################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/classes/mobile/todaydealCls.asp" -->
<%
'###############################################
'�̺�Ʈ �ű� ��Ͻ�
'###############################################
%>
<!-- #include virtual="/admin/eventmanage/common/event_function.asp"-->
<!-- #include virtual="/lib/classes/event/eventmanageCls.asp"-->
<%
Dim eCode
Dim idx , isusing , mode
Dim srcSDT , srcEDT
Dim mainStartDate, mainEndDate
Dim sDt, sTm, eDt, eTm
Dim itemurl , itemurlmo
Dim dealtitle
Dim prevDate
Dim itemid , itemname , limitno
Dim stdt , eddt , sortnum , smallImage
Dim gubun1 , gubun2 , limityn

	idx = requestCheckvar(request("idx"),16)
	srcSDT = request("sDt")
	srcEDT = request("eDt")
	prevDate = request("prevDate")

If idx = "" Then
	mode = "add"
Else
	mode = "modify"
End If

'// ������
If idx <> "" then
	dim oTodayDealOne
	set oTodayDealOne = new CMainbanner
	oTodayDealOne.FRectIdx = idx
	oTodayDealOne.GetOneContents()

	idx					=	oTodayDealOne.FOneItem.Fidx
	smallImage			=	oTodayDealOne.FOneItem.FSmallimg
	itemurl				=	oTodayDealOne.FOneItem.Fitemurl
	itemurlmo			=	oTodayDealOne.FOneItem.Fitemurlmo '2014-09-16 ����Ͽ��߰�
	dealtitle			=	oTodayDealOne.FOneItem.Fdealtitle
	mainStartDate		=	oTodayDealOne.FOneItem.Fstartdate
	mainEndDate			=	oTodayDealOne.FOneItem.Fenddate
	isusing				=	oTodayDealOne.FOneItem.Fisusing
	sortnum				=	oTodayDealOne.FOneItem.Fsortnum
	gubun1				=	oTodayDealOne.FOneItem.Fgubun1
	gubun2				=	oTodayDealOne.FOneItem.Fgubun2
	limityn				=	oTodayDealOne.FOneItem.Flimityn
	limitno				=	oTodayDealOne.FOneItem.Flimitno
	itemid				=	oTodayDealOne.FOneItem.Fitemid
	itemname			=	oTodayDealOne.FOneItem.Fitemname


	set oTodayDealOne = Nothing
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
end If

%>
<link href="/js/jqueryui/css/jquery-ui.css" rel="stylesheet">
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/jqueryui/jquery-ui-1.10.2.custom.min.js"></script>
<script type='text/javascript'>

	function getByteLength(str) {
		var p, len = 0;
		for (p = 0; p < str.length; p++) {
			(str.charCodeAt(p)  > 255) ? len += 2 : len++;
		}

		return len;
	}

	function jsSubmit(){
		var frm = document.frm;

		if (getByteLength(frm.dealtitle.value) >= 50) {
			alert("������ ª�� �Է��ϼ���(" + getByteLength(frm.dealtitle.value) + "/50)");
			return;
		}

		if (confirm('���� �Ͻðڽ��ϱ�?')){
			//frm.target = "blank";
			frm.submit();
		}
	}
	function jsgolist(){
		self.location.href="/admin/appmanage/today/todaydeal/";
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

	function onchgbox(v){
		if (v == "3"){
			$("#gubun2").css("display","block");
		}else{
			$("#gubun2").css("display","none");
		}
	}

	function fnGetItemInfo(iid) {
		$.ajax({
			type: "GET",
			url: "act_iteminfo.asp?itemid="+iid,
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
					var rst = "<img src='" + $(xml).find("itemInfo").find("item").find("smallImage").text() + "' height='80' /><br/><br/>�������� : " + $(xml).find("itemInfo").find("item").find("limitno").text() + "�� <br/>"
						rst += "��ǰ�� : <input type='text' value='" + $(xml).find("itemInfo").find("item").find("itemname").text() + "' size='40' id='tempname' onkeyup='copytxt();'/>"
					$("#lyItemInfo").fadeIn();
					$("#lyItemInfo").html(rst);
					$("#itemname").val($(xml).find("itemInfo").find("item").find("itemname").text());
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
	//�������̸�����
	function copytxt(){
		var txt1 = $("#tempname");
		var txt2 = $("#itemname");
		txt2.val(txt1.val());
	}
</script>
<table width="1000" cellpadding="2" cellspacing="1" class="a" bgcolor="#3d3d3d">
<form name="frm" method="post" action="dotodaydeal.asp">
<input type="hidden" name="mode" value="<%=mode%>">
<input type="hidden" name="adminid" value="<%=session("ssBctId")%>">
<input type="hidden" name="idx" value="<%=idx%>">
<input type="hidden" name="prevDate" value="<%=prevDate%>">
<input type="hidden" name="menupos" value="<%=menupos%>">
<input type="hidden" id="itemname" name="itemname" value="<%=itemname%>">
<tr bgcolor="#FFFFFF">
    <td bgcolor="#FFF999" align="center" width="15%">����Ⱓ</td>
    <td colspan="3">
		<input type="text" id="sDt" name="StartDate" size="10" value="<%=sDt%>" />
		<input type="text" name="sTm" size="8" value="<%=sTm%>" /> ~
		<input type="text" id="eDt" name="EndDate" size="10" value="<%=eDt%>" />
		<input type="text" name="eTm" size="8" value="<%=eTm%>" />
    </td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center" height="40">���и�</td>
	<td colspan="3">
		<div style="float:left">
		<select name="gubun1" onchange="onchgbox(this.value);" width="100">
			<option value="">=====���м���=====</option>
			<option value="1" <%=chkiif(gubun1="1","selected","")%>>TIME SALE</option>
			<option value="2" <%=chkiif(gubun1="2","selected","")%>>WISH NO.1</option>
			<option value="3" <%=chkiif(gubun1="3","selected","")%>>ISSUE ITEM</option>
		</select>&nbsp;&nbsp;
		</div>
		<div>
		<select id="gubun2" name="gubun2" style="display:<%=chkiif(gubun1="3","block","none")%>;" width="100">
			<option value="">=====�̽�����=====</option>
			<option value="1" <%=chkiif(gubun2="1","selected","")%>>���� ���԰�</option>
			<option value="2" <%=chkiif(gubun2="2","selected","")%>>HOT ITEM</option>
			<option value="3" <%=chkiif(gubun2="3","selected","")%>>SPECIAL EDITION</option>
			<option value="4" <%=chkiif(gubun2="4","selected","")%>>10x10 ONLY</option>
		</select>
		</div>
		<div style="clear:both;"></div>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center" height="40">����</td>
	<td colspan="3">
		<input type="text" name="dealtitle" size="50" value="<%=dealtitle%>"/>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center" height="40">��ǰ</td>
	<td colspan="3">
		<input type="text" name="itemid" value="<%= itemid %>" size="8" maxlength="8" class="text" require="N" onblur="fnGetItemInfo(this.value)" title="��ǰ�ڵ�" />
		<input type="button" value="��ǰ���"  >
        <div id="lyItemInfo" style="display:<%=chkIIF(itemid="","none","block")%>;">
        <%
        	if Not(itemName="" or isNull(itemName)) then
        		Response.Write "<img src='" & smallImage & "' height='80' /><br/><br/>�������� :" & limitno & "��<br/>"
	    		Response.Write "��ǰ�� : <input type='text' value='"& itemName &"' id='tempname' onkeyup='copytxt();' size='40' />"
        	end if
        %>
        </div>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center" height="40">����ϻ�ǰURL</td>
	<td colspan="3">
		<input type="text" name="itemurlmo" size="110" value="<%=itemurlmo%>"/><br/><br/>
		<span style="color:red">�� ������ǰ�� �ƴ� �Ϲ� ��ǰ�� ��쵵 URL�� �־� �ּ���<br/>
		��) http://m.10x10.co.kr/category/category_itemprd.asp?itemid=884073</span>
		<br/><br/>
		�� [ON]��ǰ���� &gt;&gt; ��ǰ�������� ��� ��ǰ �˾� Ŭ�� &gt;&gt; ��ǰ��ȣ �ϴ� URL���� &gt;&gt; ������� ��ũ ����<br/>
		��) http://m.10x10.co.kr/category/category_itemprd.asp?itemid=663507&ldv=MzIwMCAg
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center" height="40">APP��ǰURL</td>
	<td colspan="3">
		<input type="text" name="itemurl" size="110" value="<%=itemurl%>"/><br/><br/>
		<span style="color:red">�� ������ǰ�� �ƴ� �Ϲ� ��ǰ�� ��쵵 URL�� �־� �ּ���<br/>
		��) http://m.10x10.co.kr/apps/appcom/wish/web2014/category/category_itemprd.asp?itemid=884073</span>
		<br/><br/>
		�� [ON]��ǰ���� &gt;&gt; ��ǰ�������� ��� ��ǰ �˾� Ŭ�� &gt;&gt; ��ǰ��ȣ �ϴ� URL���� &gt;&gt; wishApp ��ũ ����<br/>
		��) http://m.10x10.co.kr/apps/appcom/wish/web2014/category/category_itemprd.asp?itemid=884073&ldv=OTQ4MiAg
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center">���� ��ȣ</td>
	<td colspan="3"><input type="text" name="sortnum" size="10" value="99" maxlength="3"/></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center">��뿩��</td>
	<td colspan="3"><div style="float:left;"><input type="radio" name="isusing" value="Y" <%=chkiif(isusing = "Y","checked","")%> checked />����� &nbsp;&nbsp;&nbsp; <input type="radio" name="isusing" value="N"  <%=chkiif(isusing = "N","checked","")%>/>������</div> <div style="float:right;margin-top:5px;margin-right:10px;"></div></td>
</tr>
<tr bgcolor="#FFFFFF" align="center">
    <td colspan="4"><input type="button" value=" �� �� " onClick="jsgolist();"/><input type="button" value=" �� �� " onClick="jsSubmit();"/></td>
</tr>
</form>
</table>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->

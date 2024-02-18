<%@ language=vbscript %>
<% option explicit %>
<%
'###############################################
' PageName : main_manager.asp
' Discription : ����Ʈ ���� ����
' History : 2008.04.11 ������ : �Ǽ������� ����
'			2009.04.19 �ѿ�� 2009�� �°� ����
'           2009.12.21 ������ : ���ں� �÷��� ���� ��� �߰�
'			2012.02.08 ������ : �̴ϴ޷� ��ü
'           2013.09.28 ������ : 2013������ - �߰����� �ʵ� �߰�
'           2015.04.07 ������ : 2015������ - �߰����� �ʵ� �߰�
'           2018-01-15 ����ȭ : ���� PC��� ���� �߰�
'           2018-08-30 ������ : pc, ����� ��ǰ�� ��ʿ� ȸ������, os���� �߰�
'			2019.09.27 ������ : ���Ľ����̼� �̺�ƮDB ���� ���� ����
'			2019.11.20 ������ : �̹��� ���� �߰�
'###############################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/classes/sitemasterclass/main_ContentsManageCls.asp" -->
<!-- #include virtual="/admin/eventmanage/common/event_function.asp"-->
<!-- #include virtual="/lib/classes/displaycate/displaycateCls.asp"-->
<%
dim isusing, fixtype, validdate, prevDate
dim idx, poscode, reload, gubun, edid
Dim culturecode
	idx = request("idx")
	poscode = request("poscode")
	reload = request("reload")
	gubun = request("gubun")

	isusing = request("isusing")
	fixtype = request("fixtype")
	validdate= request("validdate")
	prevDate = request("prevDate")

	culturecode = request("eC")

	if idx="" then idx=0

	if reload="on" then
			response.write "<script>opener.location.reload(); window.close();</script>"
			dbget.close()	:	response.End
	end if

	dim oMainContents
		set oMainContents = new CMainContents
		oMainContents.FRectIdx = idx
		oMainContents.GetOneMainContents

dim oposcode, defaultMapStr, defaultXMLMapStr
	set oposcode = new CMainContentsCode
	oposcode.FRectPosCode = poscode
	if poscode<>"" then
			oposcode.GetOneContentsCode

			defaultMapStr = "<map name='Map_" +oposcode.FOneItem.FPosvarName + "'>" + VbCrlf
			defaultMapStr = defaultMapStr + VbCrlf
			defaultMapStr = defaultMapStr + "</map>"

		defaultXMLMapStr = ""
			defaultXMLMapStr = defaultXMLMapStr + "<map name='Map_" +oposcode.FOneItem.FPosvarName + "'>"+ VbCrlf
			defaultXMLMapStr = defaultXMLMapStr + VbCrlf
		defaultXMLMapStr = defaultXMLMapStr + "</map>"
	end if

dim orderidx
	if oMainContents.FOneItem.forderidx = "" then
		orderidx = 99
	else
		orderidx = oMainContents.FOneItem.forderidx
		poscode = oMainContents.FOneItem.fposcode
	end if

	If gubun = "" Then
		gubun = "index"
	End If

	edid = oMainContents.FOneItem.Fworkeruserid
	If edid = "" Then
		If idx <> "" AND idx <> "0" Then
			edid = session("ssBctId")
		End If
	End If

	'// ���Ľ����̼� �ҷ�����
	Dim cultureContents, SqlStr
	Dim cultureEcode ,	cultureEtype ,cultureEname ,cultureEcomment , cultureEimagelist



	If culturecode<>"" Then
		sqlStr = "SELECT e.evt_code, d.eventtype_pc as evt_type, e.evt_name, e.evt_subcopyk, d.evt_mainimg" + vbcrlf
		sqlStr = sqlStr & " from db_event.dbo.tbl_event as e" + vbcrlf
		sqlStr = sqlStr & " LEFT JOIN [db_event].[dbo].[tbl_event_display] as d on d.evt_code=e.evt_code"
		sqlStr = sqlStr & " where e.evt_using='Y'" + vbcrlf
		sqlStr = sqlStr & " and e.evt_code="& culturecode		

		rsget.Open SqlStr, dbget, 1
		if Not rsget.Eof then
			cultureEcode		= rsget("evt_code")
			cultureEtype		= rsget("evt_type")
			cultureEname		= rsget("evt_name")
			cultureEcomment		= rsget("evt_subcopyk")
			'cultureEimagelist	= webImgUrl &"/culturestation/2009/list/" & rsget("image_list")
			cultureEimagelist	= rsget("evt_mainimg")
		end if
		rsget.close
	End If

'// Ư�� �ڵ忡 ��ũ�ؽ�Ʈ �߰�(IMG ALT �� ��)
dim IsLinkTextNeed
	IsLinkTextNeed = (InStr(",630,642,659,673,674,675,687,", ("," & poscode & ",")) > 0)

'//��ǰ�� ��� ī�װ� ����
	dim cDisp, cateIndex, cateCodeArr, cateNameArr(), cateIdx, categoryOptions
	categoryOptions = oMainContents.FOneItem.FcategoryOptions
	cateCodeArr = split(categoryOptions, ",")
	redim preserve cateNameArr(ubound(cateCodeArr))

	SET cDisp = New cDispCate
	cDisp.FCurrPage = 1
	cDisp.FPageSize = 2000
	cDisp.FRectDepth = 1
	cDisp.GetDispCateList()

	For cateIndex = 0 To cDisp.FResultCount-1
		for cateIdx = 0 to ubound(cateCodeArr) - 1
			if Cint(cDisp.FItemList(cateIndex).FCateCode) = Cint(cateCodeArr(cateIdx)) then
				cateNameArr(cateIdx) = cDisp.FItemList(cateIndex).FCateName
			end if
		next
	next
	'response.write cateNameArr
%>
<script type="text/javascript" src="/js/jquery-1.7.2.min.js"></script>
<script type="text/javascript" src="/js/jsCal/js/jscal2.js"></script>
<script type="text/javascript" src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script type="text/javascript">
<%
	'ecode ���Ľ����̼��̺�Ʈid
	'maincopy ��������
	'subcopy �߰� �ڸ�Ʈ����
	'linktext3  ���� (����)
	'xbtncolor 0/1	���м���
	'file1 �̹��� --  �̹��� �� �ְ� �����ؼ� ����ҵ�
%>
	<% if culturecode <> "" then %>
	$(function(){
		var gubuncode = "<%=cultureEtype%>";
		var frm = document.frmcontents;
			frm.ecode.value = "<%=cultureEcode%>";
			frm.maincopy.value = "<%=cultureEname%>";
			frm.subcopy.value = "<%=cultureEcomment%>";
			if (gubuncode == "0"){
				frm.xbtncolor[0].value = "0";
				frm.xbtncolor[0].checked = true;
			}else{
				frm.xbtncolor[1].value = "1";
				frm.xbtncolor[1].checked = true;
			}
			frm.linkurl.value = "/culturestation/culturestation_event.asp?evt_code=<%=cultureEcode%>";
	});
	<% end if %>

	function SaveMainContents(frm){
			if (frm.poscode.value.length<1){
					alert('������ ���� ���� �ϼ���.');
					frm.poscode.focus();
					return;
			}

			if (frm.linkurl.value.length<1 && !$("#couponRadioBtn").is(':checked') && !$("#popupBnrBtn").is(':checked')){
					alert('��ũ ���� �Է� �ϼ���.');
					frm.linkurl.focus();
					return;
			}

			if (frm.startdate.value.length!=10){
					alert('�������� �Է�  �ϼ���.');
					return;
			}

			if (frm.enddate.value.length!=10){
					alert('�������� �Է�  �ϼ���.');
					return;
			}
		<% if poscode <> "562" and poscode <> "561" then  %>
		if (!frm.altname.value){
			alert('alt���� �Է� �ϼ���.');
			frm.altname.focus();
			return;
		}
		<% end if %>

			var vstartdate = new Date(frm.startdate.value.substr(0,4), (1*frm.startdate.value.substr(5,2))-1, frm.startdate.value.substr(8,2));
			var venddate = new Date(frm.enddate.value.substr(0,4), (1*frm.enddate.value.substr(5,2))-1, frm.enddate.value.substr(8,2));

			if (vstartdate>venddate){
					alert('�������� �����Ϻ��� ������ �ȵ˴ϴ�.');
					return;
			}

			if (confirm('���� �Ͻðڽ��ϱ�?')){
					frm.submit();
			}
	}

	function ChangeLinktype(comp){
			if (comp.value=="M"){
				 document.all.link_M.style.display = "";
				 document.all.link_L.style.display = "none";
			}else{
				 document.all.link_M.style.display = "none";
				 document.all.link_L.style.display = "";
			}
	}

	//function getOnLoad(){
	//    ChangeLinktype(frmcontents.linktype.value);
	//}

	//window.onload = getOnLoad;

	function ChangeGubun(comp){
			location.href = "?gubun=<%=gubun%>&poscode=" + comp.value;
			// nothing;
	}


	function ChangeGroupGubun(comp){
			location.href = "?gubun=" + comp.value;
			// nothing;
	}

	function putLinkText(key) {
		var frm = document.frmcontents;
		switch(key) {
			case 'search':
				frm.linkurl.value='/search/search_item.asp?rect=�˻���';
				break;
			case 'event':
				frm.linkurl.value='/event/eventmain.asp?eventid=�̺�Ʈ��ȣ';
				break;
			case 'itemid':
				frm.linkurl.value='/shopping/category_prd.asp?itemid=��ǰ�ڵ�';
				break;
			case 'category':
				frm.linkurl.value='/shopping/category_list.asp?disp=ī�װ�';
				break;
			case 'brand':
				frm.linkurl.value='/street/street_brand.asp?makerid=�귣����̵�';
				break;
			case 'showbanner':
				frm.linkurl.value='/showbanner/show_view.asp?showidx=���ʾ��̵�';
				break;
			case 'culture':
				frm.linkurl.value='/culturestation/culturestation_event.asp?evt_code=�̺�Ʈ���̵�';
				break;
			case 'ground':
				frm.linkurl.value='/play/playGround.asp?idx=�׶����ȣ&contentsidx=��������ȣ';
				break;
			case 'styleplus':
				frm.linkurl.value='/play/playStylePlus.asp?idx=��Ÿ���÷�����ȣ&contentsidx=��������ȣ';
				break;
			case 'fingers':
				frm.linkurl.value='/play/playDesignFingers.asp?idx=�ΰŽ���ȣ&contentsidx=��������ȣ';
				break;
			case 'tepisode':
				frm.linkurl.value='/play/playTEpisode.asp?idx=Ƽ���Ǽҵ��ȣ&contentsidx=��������ȣ';
				break;
			case 'gift':
				frm.linkurl.value='/gift/gifttalk/';
				break;
			case 'wish':
				frm.linkurl.value='/wish/index.asp';
				break;
			case 'hitchhiker':
				frm.linkurl.value='/hitchhiker/';
				break;
			case 'giftcard':
				frm.linkurl.value='/giftcard/';
				break;
			case 'coupon':
				frm.linkurl.value='/my10x10/couponbook.asp';
				break;
		}
	}

	function cultureloadpop(){
		winLast = window.open('pop_culturelist.asp?gubun=<%=gubun%>&poscode=<%=poscode%>&pidx=<%=idx%>','pLast','width=1200,height=600, scrollbars=yes')
		winLast.focus();
	}

	function fnSelectBannerType(bannertype){
		switch (bannertype) {
			case 1 :
				$("#bnimg3").hide();
				$("#bnalt3").hide();
				$("#bnimg2").hide();
				$("#bnalt2").hide();
				$("#bnbg1").hide();
				$("#bnbg2").hide();
				$("#bnlink2").hide();
				$("#bnlink3").hide();
				break;
			case 2 :
				$("#bnbg1").show();
				$("#bnbg2").show();
				$("#bnimg2").show();
				$("#bnalt2").show();
				$("#bnlink2").show();
				$("#bnlink3").hide();
				$("#bnimg3").hide();
				$("#bnalt3").hide();
				break;
			case 3 :
				$("#bnbg1").show();
				$("#bnbg2").show();
				$("#bnimg2").show();
				$("#bnalt2").show();
				$("#bnlink2").show();
				$("#bnimg3").show();
				$("#bnalt3").show();
				$("#bnlink3").show();
				break;
		}
	}

	$(function() {
		$('input:radio[name="etctag"]').click(function(){
			if($('input:radio[name="etctag"]:checked').val()==8 || $('input:radio[name="etctag"]:checked').val()==9)
			{
				alert('�̺�Ʈ �ڵ带 �Է� ���ּ���');
				$("#saleinfo2").focus();
			}
		});
	});

var selectedCategoryArr = [];
	<%
		if categoryOptions <> "" then
			for cateIdx = 0 to ubound(cateNameArr) - 1
	%>
			var tempObj = {
				categoryCode: '<%=cateCodeArr(cateIdx)%>',
				categoryName: '<%=cateNameArr(cateIdx)%>'
			}
			selectedCategoryArr.push(tempObj);
	<%
			next
		end if
	%>
$(function(){
	dispSelectedCateNames()
})
function addCategory(){
	var cateSelectBox = document.frmcontents.categoryCode;

	var selectedObj;
	var selectedCcode, selectedCname;

	selectedCcode = cateSelectBox.value;
	selectedCname = cateSelectBox.options[cateSelectBox.selectedIndex].text.replace(" ","");

	if(chkCategory(selectedCcode))return false;

	selectedObj = {
		categoryCode: selectedCcode,
		categoryName: selectedCname
	}
	selectedCategoryArr.push(selectedObj);

	dispSelectedCateNames();
	setCategoryValues();
}
function chkCategory(selectedCcode){
	var result = false;
	selectedCategoryArr.forEach(function(item, index){
		if(item.categoryCode == selectedCcode){
			alert("�̹� �߰����ִ� ī�װ��Դϴ�.");
			result = true;
			return false;
		}
	});
	return result;
}
function dispSelectedCateNames(){

	var selectedCategoryNamesText="";

	selectedCategoryArr.forEach(function(item, index){
		selectedCategoryNamesText = selectedCategoryNamesText + "<span onclick='subCateObj("+item.categoryCode+")'>"+item.categoryName+", </span>";
	});
	$("#categoryDisplay").html(selectedCategoryNamesText);
}
function setCategoryValues(){

	var selectedCategoryCodes="";
	selectedCategoryArr.forEach(function(item, index){
		selectedCategoryCodes = selectedCategoryCodes + item.categoryCode+",";
	});
	document.frmcontents.categoryOptions.value = selectedCategoryCodes;
}
function subCateObj(selectedCode){
	selectedCategoryArr = selectedCategoryArr.filter(function(obj){
		return obj.categoryCode != selectedCode;
	});
	dispSelectedCateNames();
	setCategoryValues();
}
function chkWhiteSpace(obj){
	obj.value = obj.value.trim();
}

function fnDeleteImage(imgnum){
	$("#img"+imgnum).attr("src","");
	$("#file"+imgnum).val("");
	$("#imgurl"+imgnum).html("");
	$("#dfile"+imgnum).val("Y");
}
</script>

<table width="100%" cellpadding="2" cellspacing="1" class="a" bgcolor="#3d3d3d">
<form name="frmcontents" method="post" action="<%=uploadUrl%>/linkweb/doMainContentsRegNew.asp" onsubmit="return false;" enctype="multipart/form-data">
<input type="hidden" name="dfile1" id="dfile1">
<input type="hidden" name="dfile2" id="dfile2">
<input type="hidden" name="dfile3" id="dfile3">
	<tr bgcolor="#FFFFFF">
		<td width="150" bgcolor="#DDDDFF">Idx</td>
		<td>
			<% if oMainContents.FOneItem.Fidx<>"" then %>
				<%= oMainContents.FOneItem.Fidx %>
				<input type="hidden" name="idx" value="<%= oMainContents.FOneItem.Fidx %>">
			<% else %>
				<% '?? %>
			<% end if %>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td width="150" bgcolor="#DDDDFF">�׷챸��</td>
		<td>
			<% if oMainContents.FOneItem.Fidx<>"" then %>
				<%= oMainContents.FOneItem.Fgubun %>
				<input type="hidden" name="gubun" value="<%= oMainContents.FOneItem.Fgubun %>">
			<% else %>
				<% call DrawGroupGubunCombo("gubun", gubun, "onChange='ChangeGroupGubun(this);'") %>
			<% end if %>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td width="150" bgcolor="#DDDDFF">���и�</td>
		<td>
			<% IF gubun <> "" Then %>
				<% if oMainContents.FOneItem.Fidx<>"" then %>
					<%= oMainContents.FOneItem.Fposname %> (<%= oMainContents.FOneItem.Fposcode %>)
					<input type="hidden" name="poscode" value="<%= oMainContents.FOneItem.Fposcode %>">
				<% else %>
					<% call DrawMainPosCodeCombo("poscode", poscode, "onChange='ChangeGubun(this);'", gubun) %>
				<% end if %>
			<% Else %>
				<font color="red">�׷챸���� ���� �����ϼ���</font>
			<% End If %>
			<% If poscode = "714" Then %>
				<%'//[2018] ���Ľ����̼�%>
				&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<span><a href="" onclick="cultureloadpop();return false;">�ҷ�����</a></span>
			<% End If %>
		</td>
	</tr>
<!-- ==================================������� �߰� 2019-08-21 ======================================-->
	<% If poscode = "716" or poscode = "715"  or poscode = "708" or poscode = "733" or poscode = "734" or poscode = "735" or poscode = "738" or poscode = "739" or poscode = "730" or poscode = "729" or poscode = "728" or poscode = "707" then %>
	<script>
	$(function(){
	<% if oMainContents.FOneItem.Fbannertype = 1 then %>
		setCouponRow(1)
	<% elseif oMainContents.FOneItem.Fbannertype = 2 then%>	
		setCouponRow(2)
		setAddButton()
	<% else %>		
		setCouponRow(3)
		setAddButton()
	<% end if %>
	
	})
	function setCouponRow(v){
		if(v == 1){			
			$(".coupon-row").css("display", "none")
			$(".lyr-row").css("display", "none")			
			$(".add-btn").css("display", "none")			
		}else if(v == 2){
			$(".coupon-row").show()
			$(".lyr-row").show()
			$(".add-btn").show()
		}else{
			$(".coupon-row").css("display", "none")
			$(".lyr-row").show()
			$(".add-btn").show()
		}
	}
	function setAddButton(){
		var isChk = $('input:checkbox[id="btnFlag"]').is(':checked')
		if(isChk){
			$(".btn-row").css("display", "")
		}else{
			$(".btn-row").css("display", "none")
		}
	}
	function jsLastEvent(){	
		var winLast = window.open('pop_coupon_list.asp','pLast','width=550,height=600, scrollbars=yes')
		winLast.focus();
	}	
	</script>
	<% if poscode = "716" then %>
	<tr bgcolor="#FFFFFF">
		<td width="150" bgcolor="#DDDDFF">ä��</td>
		<td>
			<input type="radio" name="etctext" value="1" <%=ChkIIF(oMainContents.FOneItem.Fetctext = "1" or oMainContents.FOneItem.Fetctext ="", "checked", "")%>>��, ����
			<input type="radio" name="etctext" value="2" <%=ChkIIF(oMainContents.FOneItem.Fetctext = "2", "checked", "")%>>��
			<input type="radio" name="etctext" value="3" <%=ChkIIF(oMainContents.FOneItem.Fetctext = "3", "checked", "")%>>����
		</td>
	</tr>	
	<% end if %>
	<tr bgcolor="#FFFFFF">
		<td width="150" bgcolor="#DDDDFF">��ʱ���</td>
		<td>
			<input type="radio" name="bannerType" onclick="setCouponRow(1)" value="1" <%=ChkIIF(oMainContents.FOneItem.Fbannertype = 1 or oMainContents.FOneItem.Fbannertype ="", "checked", "")%>>��ũ ���
			<input type="radio" name="bannerType" onclick="setCouponRow(2)" value="2" <%=ChkIIF(oMainContents.FOneItem.Fbannertype = 2, "checked", "")%> id="couponRadioBtn">�������
			<input type="radio" name="bannerType" onclick="setCouponRow(3)" value="3" <%=ChkIIF(oMainContents.FOneItem.Fbannertype = 3, "checked", "")%> id="popupBnrBtn">�˾� ���
		</td>
	</tr>	
	<tr bgcolor="#FFFFFF" class="coupon-row" style="display:none">
		<td width="150" bgcolor="#DDDDFF">������ȣ</td>
		<td>
			<input type="number" name="couponidx" id="couponidx" value="<%=oMainContents.FOneItem.Fcouponidx%>">
			<button type="button" onclick="jsLastEvent()">����ã��</button>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF" class="lyr-row">
		<td width="150" bgcolor="#DDDDFF">���̾� �˾� ī��</td>
		<td>
			<input type="text" name="maincopy" value="<%=oMainContents.FOneItem.Fmaincopy%>" size="80" maxlength="200" class="text" /><br/>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF" class="lyr-row">
		<td width="150" bgcolor="#DDDDFF">���̾� �˾� ����ī��</td>
		<td>
			<input type="text" name="maincopy2" value="<%=oMainContents.FOneItem.Fmaincopy2%>" size="80" maxlength="200" class="text" /><br/>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF"  class="add-btn">
		<td width="150" bgcolor="#DDDDFF">��ư �߰�</td>
		<td>
			<input type="checkbox" name="etctag" id="btnFlag" onclick="setAddButton()" value="1" <%=ChkIIF(oMainContents.FOneItem.Fetctag = 1, "checked", "")%>>��ư �߰�
		</td>
	</tr>			
	<tr bgcolor="#FFFFFF" class="btn-row" style="display:none">
		<td width="150" bgcolor="#DDDDFF">��ư ī��</td>
		<td>
			<input type="text" name="subcopy" value="<%=oMainContents.FOneItem.Fsubcopy%>" size="80" maxlength="200" class="text" /><br/>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF" class="btn-row" style="display:none">
		<td width="150" bgcolor="#DDDDFF">��ư ����url</td>
		<td>
			<input type="text" name="linkurl2" value="<%=oMainContents.FOneItem.Flinkurl2%>" size="80" maxlength="200" class="text" /><br/>
		</td>
	</tr>
	<% end if %>
	<% If poscode = "707" or poscode = "708" or poscode = "715" or poscode = "716" or poscode = "725" or poscode = "728" or poscode = "732" or poscode = "734" or poscode = "739" Then %>
	<%'//[2018] ��ǰ�� ������, [2018] ��ǰ�� ���� ������, [2018] pc��ǰ�� ������(����), [2018] mo��ǰ�� ��� (����)���, ī�װ� ���� ����Ʈ ���, �̺�Ʈ �󼼹��, [2018] app ��ǰ�� ���� ������ %>
		<tr bgcolor="#FFFFFF">
			<td width="150" bgcolor="#DDDDFF">���</td>
			<td>
				<select name="targetType" class="formSlt">
					<option value="" <%=chkIIF(oMainContents.FOneItem.FtargetType="","selected","")%>>����</option>
					<option value="0" <%=chkIIF(oMainContents.FOneItem.FtargetType="0","selected","")%>>white</option>
					<option value="1" <%=chkIIF(oMainContents.FOneItem.FtargetType="1","selected","")%>>red</option>
					<option value="2" <%=chkIIF(oMainContents.FOneItem.FtargetType="2","selected","")%>>vip</option>
					<option value="3" <%=chkIIF(oMainContents.FOneItem.FtargetType="3","selected","")%>>vip gold</option>
					<option value="4" <%=chkIIF(oMainContents.FOneItem.FtargetType="4","selected","")%>>vvip</option>
				</select>
				<span>
					<span style="padding-left: 120px">ī�װ�</span>
					<select name="categoryCode" class="formSlt">
						<% For cateIndex=0 To cDisp.FResultCount-1 %>
							<option value="<%=cDisp.FItemList(cateIndex).FCateCode%>"><%=" "&cDisp.FItemList(cateIndex).FCateName%></option>
						<% next %>
					</select>
					<button type="button" onclick="addCategory();">����</button>
				</span>
				<br/>
				<span style="color:darkred">�� ī�װ��� �������� �����ø� ��ü ī�װ��� ����˴ϴ�.</span>
			</td>
		</tr>
		<tr bgcolor="#FFFFFF">
			<td width="150" bgcolor="#DDDDFF">���� ī�װ�</td>
			<td id="categoryDisplay"></td>
			<input type="hidden" name="categoryOptions" value="<%=categoryOptions%>">
		</tr>
	<% End If %>

	<% If poscode="729" or poscode="730" or poscode="733" or poscode="735" or poscode="738" Then %>
	<%'//[2018] ��ǰ�� ������(��ȸ��), [2018] ��ǰ�� ��������(��ȸ��) %>
		<tr bgcolor="#FFFFFF">
			<td width="150" bgcolor="#DDDDFF">���</td>
			<td>
				<select name="targetType" class="formSlt">
					<option value="99" selected>��ȸ��</option>
				</select>
				<span>
					<span style="padding-left: 120px">ī�װ�</span>
					<select name="categoryCode" class="formSlt">
						<% For cateIndex=0 To cDisp.FResultCount-1 %>
							<option value="<%=cDisp.FItemList(cateIndex).FCateCode%>"><%=" "&cDisp.FItemList(cateIndex).FCateName%></option>
						<% next %>
					</select>
					<button type="button" onclick="addCategory();">����</button>
				</span>
				<br/>
				<span style="color:darkred">�� ī�װ��� �������� �����ø� ��ü ī�װ��� ����˴ϴ�.</span>
			</td>
		</tr>
		<tr bgcolor="#FFFFFF">
			<td width="150" bgcolor="#DDDDFF">���� ī�װ�</td>
			<td id="categoryDisplay"></td>
			<input type="hidden" name="categoryOptions" value="<%=categoryOptions%>">
		</tr>
	<% End If %>

	<% If poscode = "708" or poscode = "716" or poscode = "739" Then %>
	<%'//[2018] ��ǰ�� ���� ������, [2018] ��ǰ�� ���� ������(�׽�Ʈ), [2018] app ��ǰ�� ���� ������%>
		<tr bgcolor="#FFFFFF">
			<td width="150" bgcolor="#DDDDFF">�ü��</td>
			<td>
				<select name="targetOS" class="formSlt">
					<option value="" <%=chkIIF(oMainContents.FOneItem.FtargetOS="","selected","")%>>��ü</option>
					<option value="I" <%=chkIIF(oMainContents.FOneItem.FtargetOS="I","selected","")%>>iOS</option>
					<option value="A" <%=chkIIF(oMainContents.FOneItem.FtargetOS="A","selected","")%>>�ȵ���̵�</option>
				</select>
			</td>
		</tr>
	<% End If %>

	<tr bgcolor="#FFFFFF">
		<td width="150" bgcolor="#DDDDFF">��ũ����</td>
		<td>
			<% IF gubun <> "" Then %>
				<% if oMainContents.FOneItem.Fidx<>"" then %>
					<%= oMainContents.FOneItem.getlinktypeName %>
					<input type="hidden" name="linktype" value="<%= oMainContents.FOneItem.Flinktype %>">
				<% else %>
					<% if poscode<>"" then %>
						<%= oposcode.FOneItem.getlinktypeName %>
						<input type="hidden" name="linktype" value="<%= oposcode.FOneItem.Flinktype %>">
					<% else %>
						<font color="red">������ ���� �����ϼ���</font>
					<% end if %>
				<% end if %>
			<% Else %>
				<font color="red">�׷챸���� ���� �����ϼ���</font>
			<% End If %>
		</td>
	</tr>

	<tr bgcolor="#FFFFFF">
		<td width="150" bgcolor="#DDDDFF">���뱸��(�ݿ��ֱ�)</td>
		<td>
			<% IF gubun <> "" Then %>
				<% if oMainContents.FOneItem.Fidx<>"" then %>
					<%= oMainContents.FOneItem.getfixtypeName %>
					<input type="hidden" name="fixtype" value="<%= oMainContents.FOneItem.Ffixtype %>">
				<% else %>
					<% if poscode<>"" then %>
						<%= oposcode.FOneItem.getfixtypeName %>
						<input type="hidden" name="fixtype" value="<%= oposcode.FOneItem.Ffixtype %>">
					<% else %>
						<font color="red">������ ���� �����ϼ���</font>
					<% end if %>
				<% end if %>
			<% Else %>
				<font color="red">�׷챸���� ���� �����ϼ���</font>
			<% End If %>
		</td>
	</tr>

	<tr bgcolor="#FFFFFF">
		<td width="150" bgcolor="#DDDDFF">�켱����</td>
		<td>
			<% IF gubun <> "" Then %>
				<% if oMainContents.FOneItem.Fidx<>"" then %>
					<input type="text" name="orderidx" size=5 value="<%= orderidx %>" class="text" />
				<% else %>
					<% if poscode<>"" then %>
						<input type="text" name="orderidx" size=5 value="<%= orderidx %>" class="text" />
					<% else %>
						<font color="red">������ ���� �����ϼ���</font>
					<% end if %>
				<% end if %>
			<% Else %>
				<font color="red">�׷챸���� ���� �����ϼ���</font>
			<% End If %>
		</td>
	</tr>


	<% If poscode = "727" Then %>
	<%'// �˻� ��� ������ ���(������) %>
		<tr bgcolor="#FFFFFF">
			<td width="150" bgcolor="#DDDDFF">�˻� Ű����(�ʼ�)</td>
			<td><textarea name="itemDesc" class="textarea" style="width:100%;height:80px;"><%= oMainContents.FOneItem.FitemDesc %></textarea></td>
		</tr>
	<% Else %>
		<tr bgcolor="#FFFFFF">
			<td width="150" bgcolor="#DDDDFF">�۾� ��û����</td>
			<td><textarea name="itemDesc" class="textarea" style="width:100%;height:80px;"><%= oMainContents.FOneItem.FitemDesc %></textarea></td>
		</tr>
	<% End If %>



	<% If poscode = "706" or poscode="720" or poscode="722" or poscode="723" or poscode="724" or poscode="731" Then %>
	<%'// [2018] PC ��� �ֻ�� ����, [2018] ���κ��̺�Ʈ���1~3, ���������� ��ܹ��, ���������� �ϴܹ��, �α��ι��, ����� �α��� ��� %>
		<tr bgcolor="#FFFFFF">
			<td width="150" bgcolor="#DDDDFF">��� Ÿ��</td>
			<td>
				<input type="radio" name="bannertype" value="1"<% If oMainContents.FOneItem.Fbannertype="1" Or oMainContents.FOneItem.Fbannertype="" Then Response.write " checked" %> onclick="fnSelectBannerType(1);">1��&nbsp;&nbsp;
				<input type="radio" name="bannertype" value="2"<% If oMainContents.FOneItem.Fbannertype="2" Then Response.write " checked" %> onclick="fnSelectBannerType(2);">2��&nbsp;&nbsp;
				<%'// ���������� ��,�ϴ� ���, �α��� ��ʴ� 3�� ������� �ʴ´�. %>
				<% If not(poscode="722" or poscode="723" or poscode="724" or poscode="731") Then %>
					<input type="radio" name="bannertype" value="3"<% If oMainContents.FOneItem.Fbannertype="3" Then Response.write " checked" %> onclick="fnSelectBannerType(3);">3��
				<% End If %>
			</td>
		</tr>
	<% End If %>

	<%
		'��ũ �ؽ�Ʈ ���� Ȯ��
		dim chkText: chkText="N"
		IF gubun<>"" Then
			if oMainContents.FOneItem.Fidx<>"" then
				if oMainContents.FOneItem.FLinkType="T" then
					chkText="Y"
				End If
			elseif poscode<>"" then
				if oposcode.FOneItem.FLinkType="T" then
					chkText="Y"
				End If
			end if
		end if

		'2013/09/28 ������ �߰� poscode ���
		If oMainContents.FResultCount > 0 Then
			Dim oSQL
			oSQL = " SELECT poscode FROM [db_sitemaster].[dbo].tbl_main_contents where idx = '"&oMainContents.FOneItem.Fidx&"'  "
			rsget.open oSQL, dbget, 1
			poscode = rsget("poscode")
			rsget.close
		End If
	%>

	<% IF chkText="Y" or (IsLinkTextNeed = True) then %>
		<tr bgcolor="#FFFFFF">
			<td width="150" bgcolor="#DDDDFF"><%=chkIIF(poscode="630" or poscode="687","����","��ũ �ؽ�Ʈ")%></td>
			<td><input type="text" name="linkText" value="<%= oMainContents.FOneItem.FlinkText %>" size="32" maxlength="64" class="text" /> </td>
		</tr>

		<% if poscode="630" or poscode="687" then %>
			<tr bgcolor="#FFFFFF">
				<td width="150" bgcolor="#DDDDFF">�ٹ����� �ΰ� ����</td>
				<td>
					<label><input type="radio" name="linkText2" value="wht" <%=chkIIF(oMainContents.FOneItem.FlinkText2="wht" or oMainContents.FOneItem.FlinkText2="","checked","")%> />ȭ��Ʈ</label>
					<label><input type="radio" name="linkText2" value="red" <%=chkIIF(oMainContents.FOneItem.FlinkText2="red","checked","")%> />����</label>
				</td>
			</tr>
			<tr bgcolor="#FFFFFF">
				<td width="150" bgcolor="#DDDDFF">��� ����</td>
				<td>
					<label><input type="radio" name="linkText4" value="fix" <%=chkIIF(oMainContents.FOneItem.FlinkText4="fix" or oMainContents.FOneItem.FlinkText4="","checked","")%> />FIXED</label>
					<label><input type="radio" name="linkText4" value="wide" <%=chkIIF(oMainContents.FOneItem.FlinkText4="wide","checked","")%> />WIDE</label>
				</td>
			</tr>
		<% else %>
			<tr bgcolor="#FFFFFF">
				<td width="150" bgcolor="#DDDDFF">�߰� �ؽ�Ʈ #1 (����)</td>
				<td><input type="text" name="linkText2" value="<%= oMainContents.FOneItem.FlinkText2 %>" size="40" maxlength="128" class="text" /> </td>
			</tr>
			<tr bgcolor="#FFFFFF">
				<td width="150" bgcolor="#DDDDFF">�߰� �ؽ�Ʈ #2 (����)</td>
				<td><textarea name="linkText3" style="width:100%;height:60px;" class="text"><%= oMainContents.FOneItem.FlinkText3 %></textarea></td>
			</tr>
		<% end if %>
	<% end if %>

	<% if chkText<>"Y" then %>
		<% If poscode="688" Then %>
		<%'// [2015]������ %>
			<tr bgcolor="#FFFFFF">
				<td width="150" bgcolor="#DDDDFF">��� Ÿ��Ʋ(bold)</td>
				<td><input type="text" name="linkText2" value="<%= oMainContents.FOneItem.FlinkText2 %>" size="40" maxlength="128" class="text" /> </td>
			</tr>
			<tr bgcolor="#FFFFFF">
				<td width="150" bgcolor="#DDDDFF">�ϴ� ��ǰ����</td>
				<td><input type="text" name="linkText3" value="<%= oMainContents.FOneItem.FlinkText3 %>" size="40" maxlength="128" class="text" /></td>
			</tr>
			<tr bgcolor="#FFFFFF">
				<td width="150" bgcolor="#DDDDFF">������</td>
				<td><input type="text" name="linkText4" value="<%= oMainContents.FOneItem.FlinkText4 %>" size="40" maxlength="128" class="text" />
					<br>�� ������ �ۼ��� �ϴ� ��ǰ������ �������� ����
				</td>
			</tr>
		<% End If %>

		<% If poscode="689" Then %>
		<%'// [2015]JUST1DAY or �ָ�Ư�� %>
			<tr bgcolor="#FFFFFF">
				<td width="150" bgcolor="#DDDDFF">Ÿ��Ʋ��</td>
				<td><input type="text" name="linkText2" value="<%= oMainContents.FOneItem.FlinkText2 %>" size="40" maxlength="128" class="text" />
					<br />�� �Է� ���ϸ� �⺻���� Just1Day�� �ָ�Ư�� ����<br/>�� ����Ư�� �� �Է��ϸ� ������ ������ ����Ư�� ���ڰ� ��µ�.
				</td>
			</tr>
			<tr bgcolor="#FFFFFF">
				<td width="150" bgcolor="#DDDDFF">�󼼼���</td>
				<td><input type="text" name="linkText3" value="<%= oMainContents.FOneItem.FlinkText3 %>" size="40" maxlength="128" class="text" /></td>
			</tr>
		<% End If %>

		<% If poscode="690" Or poscode="691" Or poscode="692" Or poscode="693" Or poscode="699" Then %>
		<% '// [2015]��ܹ��2��#1, [2015]��ܹ��2��#2, [2015]��ܹ��2��#3, [2015]��Ƽ���ʹ�� %>
			<tr bgcolor="#FFFFFF">
				<td width="150" bgcolor="#DDDDFF">��� Ÿ��Ʋ(bold)</td>
				<td><input type="text" name="linkText2" value="<%= oMainContents.FOneItem.FlinkText2 %>" size="40" maxlength="128" class="text" /> </td>
			</tr>
			<tr bgcolor="#FFFFFF">
				<td width="150" bgcolor="#DDDDFF">�ϴ� ��ǰ����</td>
				<td><input type="text" name="linkText3" value="<%= oMainContents.FOneItem.FlinkText3 %>" size="40" maxlength="128" class="text" /></td>
			</tr>
		<% End If %>

		<% If poscode = "710" Then %>
		<%'// 2018 ���� �Ѹ� %>
			<tr bgcolor="#FFFFFF">
				<td width="150" bgcolor="#DDDDFF">����</td>
				<td>
					�� : # <input type="text" name="linkText" value="<%= oMainContents.FOneItem.FlinkText %>" size="20" maxlength="6" class="text" /><br/>
					�� : # <input type="text" name="bgcode" value="<%=oMainContents.FOneItem.Fbgcode%>" size="20" maxlength="6" class="text">
				</td>
			</tr>
			<tr bgcolor="#FFFFFF">
				<td width="150" bgcolor="#DDDDFF">��� ����</td>
				<td>
					<label><input type="radio" name="linkText4" value="fix" <%=chkIIF(oMainContents.FOneItem.FlinkText4="fix" or oMainContents.FOneItem.FlinkText4="","checked","")%> />FIXED</label>
					<label><input type="radio" name="linkText4" value="wide" <%=chkIIF(oMainContents.FOneItem.FlinkText4="wide","checked","")%> />WIDE</label>
				</td>
			</tr>
			<tr bgcolor="#FFFFFF">
				<td width="150" bgcolor="#DDDDFF">��Ʈ�÷�����</td>
				<td>
					<input type="radio" name="xbtncolor" value="0" <%=chkiif(oMainContents.FOneItem.Fxbtncolor = "" Or oMainContents.FOneItem.Fxbtncolor = "0" ,"checked","")%>/> : black
					<input type="radio" name="xbtncolor" value="1" <%=chkiif(oMainContents.FOneItem.Fxbtncolor = "1" ,"checked","")%>/> : white
				</td>
			</tr>
			<tr bgcolor="#FFFFFF">
				<td width="150" bgcolor="#DDDDFF">����ī��</td>
				<td>
					<input type="text" name="maincopy" value="<%=oMainContents.FOneItem.Fmaincopy%>" size="50" maxlength="25" class="text" /><br/>
					<input type="text" name="maincopy2" value="<%=oMainContents.FOneItem.Fmaincopy2%>" size="80" maxlength="60" class="text" />
				</td>
			</tr>
			<tr bgcolor="#FFFFFF">
				<td width="150" bgcolor="#DDDDFF">����ī��</td>
				<td><input type="text" name="subcopy" value="<%=oMainContents.FOneItem.Fsubcopy%>" size="80" maxlength="50" class="text" /></td>
			</tr>
			<tr bgcolor="#FFFFFF">
				<td width="150" bgcolor="#DDDDFF">�±�</td>
				<td>
					<input type="radio" name="etctag" value="1" <%=chkiif(oMainContents.FOneItem.Fetctag="1" Or oMainContents.FOneItem.Fetctag="","checked","")%>> ����
					<input type="radio" name="etctag" value="2" <%=chkiif(oMainContents.FOneItem.Fetctag="2","checked","")%>> ����
					<input type="radio" name="etctag" value="3" <%=chkiif(oMainContents.FOneItem.Fetctag="3","checked","")%>> ���� <br/>
					<input type="radio" name="etctag" value="4" <%=chkiif(oMainContents.FOneItem.Fetctag="4","checked","")%>> GIFT
					<input type="radio" name="etctag" value="5" <%=chkiif(oMainContents.FOneItem.Fetctag="5","checked","")%>> 1+1
					<input type="radio" name="etctag" value="6" <%=chkiif(oMainContents.FOneItem.Fetctag="6","checked","")%>> ��Ī
					<input type="radio" name="etctag" value="7" <%=chkiif(oMainContents.FOneItem.Fetctag="7","checked","")%>> ����
					<input type="radio" name="etctag" value="8" <%=chkiif(oMainContents.FOneItem.Fetctag="8","checked","")%>> ������ �ڵ� ����(AŸ��-�߾�)
					<input type="radio" name="etctag" value="9" <%=chkiif(oMainContents.FOneItem.Fetctag="9","checked","")%>> ������ �ڵ� ����(����)
					<input type="radio" name="etctag" value="10" <%=chkiif(oMainContents.FOneItem.Fetctag="10","checked","")%>> ������ �ڵ� ����(BŸ��-�ű�)
					 <br/>
					�� �Ѱ����� ���� �ϼ���.<br/><br/>
					<input type="checkbox" name="tag_only" value="Y" <%=chkiif(oMainContents.FOneItem.Ftag_only="Y","checked","")%>> �ܵ�<br/><br/>
					<input type="text" name="etctext" value="<%=oMainContents.FOneItem.Fetctext%>" size="20" maxlength="30" class="text" />�� ����,���� �ϰ�츸 �Է� �ϼ���<br/>
				</td>
			</tr>
			<tr bgcolor="#FFFFFF">
				<td width="150" bgcolor="#DDDDFF">�̺�Ʈ �ڵ�</td>
				<td>
					<span><input type="text" id="saleinfo2" name="evt_code" value="<%= oMainContents.FOneItem.FEvt_Code %>" size="20" maxlength="10" class="text" /></span>
					<p class="tPad05"><span class="rMar10"><strong>�� ������ �ڵ� ���� �� �̺�Ʈ ���� üũ (����� , ����) ���� X ��</strong></span></p>
				</td>
			</tr>
		<% End If %>

		<% If poscode="714" Then %>
		<%'// 2018 ���Ľ����̼�%>
			<input type="hidden" name="ecode" value=""/><%' cultureidx %>
			<tr bgcolor="#FFFFFF">
				<td width="150" bgcolor="#DDDDFF">����ī��</td>
				<td><input type="text" name="maincopy" value="<%=oMainContents.FOneItem.Fmaincopy%>" size="50" maxlength="25" class="text" /></td>
			</tr>
			<tr bgcolor="#FFFFFF">
				<td width="150" bgcolor="#DDDDFF">����ī��</td>
				<td><input type="text" name="subcopy" value="<%=oMainContents.FOneItem.Fsubcopy%>" size="80" maxlength="60" class="text" /></td>
			</tr>
			<tr bgcolor="#FFFFFF">
				<td width="150" bgcolor="#DDDDFF">����</td>
				<td><textarea name="linkText3" style="width:100%;height:60px;" class="text"><%= oMainContents.FOneItem.FlinkText3 %></textarea></td>
			</tr>
			<tr bgcolor="#FFFFFF">
				<td width="150" bgcolor="#DDDDFF">���м���</td>
				<td>
					<input type="radio" name="xbtncolor" value="0" <%=chkiif(oMainContents.FOneItem.Fxbtncolor = "" Or oMainContents.FOneItem.Fxbtncolor = "0" ,"checked","")%>/> : ������
					<input type="radio" name="xbtncolor" value="1" <%=chkiif(oMainContents.FOneItem.Fxbtncolor = "1" ,"checked","")%>/> : �о��
				</td>
			</tr>
		<% End If %>

		<tr bgcolor="#FFFFFF">
			<td width="150" bgcolor="#DDDDFF">�̹���1</td>
			<td>
				<% If poscode <> "714" Then %>
				<%'// 2018 ���Ľ����̼��� �ƴҰ�츸 %>
					<input type="file" name="file1" value="" id="file1" size="32" maxlength="32" class="file">
				<% End If %>

				<% if oMainContents.FOneItem.GetImageUrl<>"" then %>
					<br>
					<img src="<%= oMainContents.FOneItem.GetImageUrl %>" id="img1" style="max-width:600px;" />
					<br><span id="imgurl1"> <%= oMainContents.FOneItem.GetImageUrl %>&nbsp;&nbsp;<input type="button" value=" ���� " onClick="fnDeleteImage('1');"></span>
				<% end if %>

				<% '���Ľ����̼� %>
				<% If oMainContents.FOneItem.Fidx = "" And poscode = "714" Then %>
					<br>
					<img src="<%=cultureEimagelist %>" style="max-width:600px;" />
					<br> <%= cultureEimagelist %> <br/><br/> �� �̹��� ������ ���Ľ����̼� ���ο��� ���ּ���
				<% ElseIf oMainContents.FOneItem.Fidx <> "" And poscode = "714" Then %>
					<br>
					<img src="<%=oMainContents.FOneItem.Fcultureimage %>" style="max-width:600px;" />
					<br> <%=oMainContents.FOneItem.Fcultureimage %> <br/><br/> �� �̹��� ������ ���Ľ����̼� ���ο��� ���ּ���
				<% End If %>
			</td>
		</tr>

		<tr bgcolor="#FFFFFF">
			<td width="150" bgcolor="#DDDDFF">��Ʈ��1 (�ʼ�)</td>
			<td><input type="text" name="altname" value="<%=oMainContents.FOneItem.Faltname%>" size="20" maxlength="20"> </td>
		</tr>

		<% If poscode = "706" or poscode = "720" or poscode="722" or poscode="723" or poscode="724" or poscode="731" Then %>
		<%'// [2018] PC ��� �ֻ�� ����, [2018] ���κ��̺�Ʈ���1~3, ���������� ��ܹ��, ���������� �ϴܹ��, �α��ι�� %>
			<tr bgcolor="#FFFFFF" id="bnimg2" style="display:<% If oMainContents.FOneItem.Fbannertype="1" or oMainContents.FOneItem.Fbannertype="" Then Response.write "none" %>">
				<td width="150" bgcolor="#DDDDFF">�̹���2</td>
				<td>
					<input type="file" name="file2" id="file2" value="" size="32" maxlength="32" class="file">
					<% if oMainContents.FOneItem.GetImageUrl2<>"" then %>
						<br>
						<img src="<%= oMainContents.FOneItem.GetImageUrl2 %>" id="img2" style="max-width:600px;" />
						<br> <span id="imgurl2"> <%= oMainContents.FOneItem.GetImageUrl2 %>&nbsp;&nbsp;<input type="button" value=" ���� " onClick="fnDeleteImage('2');"></span>
					<% end if %>
				</td>
			</tr>
			<tr bgcolor="#FFFFFF" id="bnalt2" style="display:<% If oMainContents.FOneItem.Fbannertype="1" or oMainContents.FOneItem.Fbannertype="" Then Response.write "none" %>">
				<td width="150" bgcolor="#DDDDFF">��Ʈ��2 (�ʼ�)</td>
				<td><input type="text" name="altname2" value="<%=oMainContents.FOneItem.Faltname2%>" size="20" maxlength="20"> </td>
			</tr>
			<tr bgcolor="#FFFFFF" id="bnimg3" style="display:<% If oMainContents.FOneItem.Fbannertype<>"3" Or oMainContents.FOneItem.Fbannertype="" Then Response.write "none" %>">
				<td width="150" bgcolor="#DDDDFF">�̹���3</td>
				<td>
					<input type="file" name="file3" id="file3" value="" size="32" maxlength="32" class="file">
					<% if oMainContents.FOneItem.GetImageUrl3<>"" then %>
						<br>
						<img src="<%= oMainContents.FOneItem.GetImageUrl3 %>" id="img3" style="max-width:600px;" />
						<br> <span id="imgurl3"> <%= oMainContents.FOneItem.GetImageUrl3 %>&nbsp;&nbsp;<input type="button" value=" ���� " onClick="fnDeleteImage('3');"></span>
					<% end if %>
				</td>
			</tr>
			<tr bgcolor="#FFFFFF" id="bnalt3" style="display:<% If oMainContents.FOneItem.Fbannertype<>"3" Or oMainContents.FOneItem.Fbannertype="" Then Response.write "none" %>">
				<td width="150" bgcolor="#DDDDFF">��Ʈ��3 (�ʼ�)</td>
				<td><input type="text" name="altname3" value="<%=oMainContents.FOneItem.Faltname3%>" size="20" maxlength="20"> </td>
			</tr>
		<% End If %>

		<% If gubun <> "PCbanner" and gubun <> "MAbanner" And poscode <> "706" And poscode <> "720" And poscode<>"722" And poscode<>"723" And poscode<>"724" And poscode<>"736" Then %>
		<%'// pc���, ����Ͼ۹��, ��� �ֻ��, ���κ��̺�Ʈ�� �ƴҰ�� %>
			<tr bgcolor="#FFFFFF">
				<% If poscode = "721" then %>
					<% '// ���� �÷��� ����� ��츸 Ÿ��Ʋ�� ���� %>
					<td width="150" bgcolor="#DDDDFF">���콺 ���� �� �̹���</td>
				<% Else %>
					<td width="150" bgcolor="#DDDDFF">�̹��� (����)</td>
				<% End If %>
				<td>
					<input type="file" name="file2" id="file2" value="" size="32" maxlength="32" class="file">
					<% if oMainContents.FOneItem.GetImageUrl2<>"" then %>
						<br>
						<img src="<%= oMainContents.FOneItem.GetImageUrl2 %>" id="img2" style="max-width:600px;" />
						<br> <span id="imgurl2"> <%= oMainContents.FOneItem.GetImageUrl2 %>&nbsp;&nbsp;<input type="button" value=" ���� " onClick="fnDeleteImage('2');"></span>
					<% end if %>
				</td>
			</tr>
		<% End If %>

		<tr bgcolor="#FFFFFF">
			<td width="150" bgcolor="#DDDDFF">�̹���Width</td>
			<td>
				<% IF gubun <> "" Then %>
					<% if oMainContents.FOneItem.Fidx<>"" then %>
						<input type="text" name="imagewidth" value="<%= oMainContents.FOneItem.Fimagewidth %>" size="8" maxlength="16">
						<% If poscode="720" Then %>
						<%'// ���κ��̺�Ʈ %>
							(�̹��� 1�� ����)
						<% End If %>
					<% else %>
						<% if poscode<>"" then %>
							<%= oposcode.FOneItem.Fimagewidth %>
							<% If poscode="720" Then %>
							<%'// ���κ��̺�Ʈ %>
								(�̹��� 1�� ����)
							<% End If %>
						<% else %>
							<font color="red">������ ���� �����ϼ���</font>
						<% end if %>
					<% end if %>
				<% Else %>
					<font color="red">�׷챸���� ���� �����ϼ���</font>
				<% End If %>
			</td>
		</tr>
		<tr bgcolor="#FFFFFF">
			<td width="150" bgcolor="#DDDDFF">�̹���Height</td>
			<td>
				<% IF gubun <> "" Then %>
					<% if oMainContents.FOneItem.Fidx<>"" then %>
						<input type="text" name="imageheight" value="<%= oMainContents.FOneItem.Fimageheight %>" size="8" maxlength="16">
					<% else %>
						<% if poscode<>"" then %>
							<%= oposcode.FOneItem.Fimageheight %>
						<% else %>
							<font color="red">������ ���� �����ϼ���</font>
						<% end if %>
					<% end if %>
				<% Else %>
					<font color="red">�׷챸���� ���� �����ϼ���</font>
				<% End If %>
			</td>
		</tr>
	<% End If %>

	<tr bgcolor="#FFFFFF">
		<td width="150" bgcolor="#DDDDFF">��ũ��1</td>
		<td>
			<% IF gubun <> "" Then %>
				<% if oMainContents.FOneItem.Fidx<>"" then %>
					<% if oMainContents.FOneItem.FLinkType="M" then %>
						<textarea name="linkurl" style="width:100%;height:120px;"><%= oMainContents.FOneItem.Flinkurl %></textarea>
					<% else %>
						<% if oMainContents.FOneItem.Fposcode = 539 Then%>
							<textarea name="linkurl" style="width:100%;height:120px;"><%= oMainContents.FOneItem.Flinkurl %></textarea>
						<% Else%>
							<input type="text" name="linkurl" value="<%= oMainContents.FOneItem.Flinkurl %>" maxlength="128" style="width:100%;" class="text">
						<% End If %>
					<% end if %>
				<% else %>
					<% if poscode<>"" then %>
						<% if oposcode.FOneItem.FLinkType="M" then %>
							<textarea name="linkurl" style="width:100%;height:120px;"><%= defaultMapStr %></textarea>
							<br>(�̹����� ������ ���� ����)
						<% elseif oposcode.FOneItem.FLinkType="B" then %>
							<input type="text" class="text_ro" name="linkurl" value="/" maxlength="128" size="40" readonly>
						<% elseif poscode="539" Then %>
							<textarea name="linkurl" style="width:100%;height:120px;"><%= defaultXMLMapStr %></textarea>
							<br>(�̹����� ������ ���� ����, href���Ͽ� ��ũ�־��ּ���)
						<% Else %>
							<input type="text" name="linkurl" value="" maxlength="128" style="width:100%;" class="text">
							<br>ex)<br/>
							- <span style="cursor:pointer" onClick="putLinkText('event');">�̺�Ʈ ��ũ : /event/eventmain.asp?eventid=<span style="color:darkred">�̺�Ʈ�ڵ�</span></span><br/>
							- <span style="cursor:pointer" onClick="putLinkText('itemid');">��ǰ�ڵ� ��ũ : /shopping/category_prd.asp?itemid=<span style="color:darkred">��ǰ�ڵ� (O)</span></span><br/>
							- <span style="cursor:pointer" onClick="putLinkText('category');">ī�װ� ��ũ : /shopping/category_list.asp?disp=<span style="color:darkred">ī�װ�</span></span><br/>
							- <span style="cursor:pointer" onClick="putLinkText('brand');">�귣����̵� ��ũ : /street/street_brand.asp?makerid=<span style="color:darkred">�귣����̵�</span></span><br/>
							- <span style="cursor:pointer" onClick="putLinkText('hitchhiker');">��ġ����Ŀ ��ũ : /hitchhiker/</span><br/>
							- <span style="cursor:pointer" onClick="putLinkText('giftcard');">����Ʈī�� ��ũ : /giftcard/</span><br/>
							- <span style="cursor:pointer" onClick="putLinkText('culture');">���Ľ����̼� ��ũ : /culturestation/culturestation_event.asp?evt_code=<span style="color:darkred">�̺�Ʈ���̵�</span></span><br/>
							- <span style="cursor:pointer" onClick="putLinkText('coupon');">������ ��ũ : /my10x10/couponbook.asp
						<% end if %>
					<% else %>
						<font color="red">������ ���� �����ϼ���</font>
					<% end if %>
				<% end if %>
			<% Else %>
				<font color="red">�׷챸���� ���� �����ϼ���</font>
			<% End If %>
		</td>
	</tr>

	<% If poscode = "706" or poscode="720" or poscode="722" or poscode="723" or poscode="724" or poscode="731" Then %>
	<%'// [2018] PC ��� �ֻ�� ����, [2018] ���κ��̺�Ʈ���1~3, ���������� ��, �ϴ�, �α��� ��� %>
		<tr bgcolor="#FFFFFF" id="bnlink2" style="display:<% If oMainContents.FOneItem.Fbannertype="1" Or oMainContents.FOneItem.Fbannertype="" Then Response.write "none" %>">
			<td width="150" bgcolor="#DDDDFF">��ũ��2</td>
			<td>
				<input type="text" name="linkurl2" value="<%= oMainContents.FOneItem.Flinkurl2 %>" maxlength="128" style="width:100%;" class="text">
			</td>
		</tr>
		<tr bgcolor="#FFFFFF" id="bnlink3" style="display:<% If oMainContents.FOneItem.Fbannertype<>"3" Or oMainContents.FOneItem.Fbannertype="" Then Response.write "none" %>">
			<td width="150" bgcolor="#DDDDFF">��ũ��3</td>
			<td>
				<input type="text" name="linkurl3" value="<%= oMainContents.FOneItem.Flinkurl3 %>" maxlength="128" style="width:100%;" class="text">
			</td>
		</tr>

		<% If not(poscode = "720" or poscode="722" or poscode="723" or poscode="724" or poscode="731") Then %>
		<%'// ���κ��̺�Ʈ, ���������� ��, �ϴ�, �α��� ��� ������ %>
			<tr bgcolor="#FFFFFF">
				<td width="150" bgcolor="#DDDDFF">�¿��� BG�÷��ڵ�</td>
				<td>
					<span  id="bnbg1" style="display:<% If oMainContents.FOneItem.Fbannertype="1" Or oMainContents.FOneItem.Fbannertype="" Then Response.write "none" %>">�� : </span>#<input type="text" name="bgcode" value="<%=oMainContents.FOneItem.Fbgcode%>" size="20" maxlength="6">
					<div  id="bnbg2" style="display:<% If oMainContents.FOneItem.Fbannertype="1" Or oMainContents.FOneItem.Fbannertype="" Then Response.write "none" %>">�� : #<input type="text" name="bgcode2" value="<%=oMainContents.FOneItem.Fbgcode2%>" size="20" maxlength="6"></div>
				</td>
			</tr>
			<tr bgcolor="#FFFFFF">
				<td width="150" bgcolor="#DDDDFF">X��ư����</td>
				<td>
					<input type="radio" name="xbtncolor" value="0" <%=chkiif(oMainContents.FOneItem.Fxbtncolor = "" Or oMainContents.FOneItem.Fxbtncolor = "0" ,"checked","")%>/> : ȭ��Ʈ
					<input type="radio" name="xbtncolor" value="1" <%=chkiif(oMainContents.FOneItem.Fxbtncolor = "1" ,"checked","")%>/> : black
				</td>
			</tr>
		<% End If %>
	<% End If %>

	<tr bgcolor="#FFFFFF">
		<td width="150" bgcolor="#DDDDFF">�ݿ�������</td>
		<td>
			<input id="startdate" name="startdate" value="<%=chkiif(idx=0,prevDate,Left(oMainContents.FOneItem.Fstartdate,10))%>" class="text" size="10" maxlength="10" />
			<img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="startdate_trigger" border="0" style="cursor:pointer;" align="absbottom" />
			<% if oMainContents.FOneItem.Ffixtype="R" or poscode="687" Or gubun = "PCbanner" Or gubun = "MAbanner" then %> <!-- �ǽð��ΰ�� / �� �ϴ����� ���� (���߿� �ð������� ������ False ����)-->
				<input type="text" name="startdatetime" size="2" maxlength="2" value="<%= Format00(2,Hour(oMainContents.FOneItem.Fstartdate)) %>" />(�� 00~23)
				<input type="text" name="dummy0" value="00:00" size="6" readonly class="text_ro" />
			<% else %>
				<input type="text" name="dummy0" value="00:00:00" size="8" readonly class="text_ro" />
			<% end if %>
			<script type="text/javascript">
				var CAL_Start = new Calendar({
					inputField : "startdate",
					trigger    : "startdate_trigger",
					onSelect: function() {
						var date = Calendar.intToDate(this.selection.get());
						CAL_End.args.min = date;
						CAL_End.redraw();
						this.hide();
					},
					bottomBar: true,
					dateFormat: "%Y-%m-%d"
				});
			</script>
		</td>
	</tr>

	<tr bgcolor="#FFFFFF">
		<td width="150" bgcolor="#DDDDFF">�ݿ�������</td>
		<td>
			<input id="enddate" name="enddate" value="<%=chkiif(idx=0,prevDate,Left(oMainContents.FOneItem.Fenddate,10)) %>" class="text" size="10" maxlength="10" />
			<img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="enddate_trigger" border="0" style="cursor:pointer" align="absbottom" />
			<% if oMainContents.FOneItem.Ffixtype="R"  or poscode="687" Or gubun = "PCbanner" Or gubun = "MAbanner" then %> <!-- �ǽð��ΰ�� -->
				<input type="text" name="enddatetime" size="2" maxlength="2" value="<%= ChkIIF(oMainContents.FOneItem.Fenddate="","23",Format00(2,Hour(oMainContents.FOneItem.Fenddate))) %>">(�� 00~23)
				<input type="text" name="dummy1" value="59:59" size="6" readonly class="text_ro" />
			<% else %>
				<input type="text" name="dummy1" value="23:59:59" size="8" readonly class="text_ro" />
			<% end if %>
			<script type="text/javascript">
				var CAL_End = new Calendar({
					inputField : "enddate",
					trigger    : "enddate_trigger",
					onSelect: function() {
						var date = Calendar.intToDate(this.selection.get());
						CAL_Start.args.max = date;
						CAL_Start.redraw();
						this.hide();
					},
					bottomBar: true,
					dateFormat: "%Y-%m-%d"
				});
			</script>
		</td>
	</tr>

	<tr bgcolor="#FFFFFF">
		<td width="150" bgcolor="#DDDDFF">�����</td>
		<td>
			<%= oMainContents.FOneItem.Fregdate %> (<%= oMainContents.FOneItem.Fregname %>)
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td width="150" bgcolor="#DDDDFF">�۾���</td>
		<td>
			<% If idx <> "" AND idx <> "0" Then %>
				���� �۾��� : <%=oMainContents.FOneItem.Fworkername%><input type="hidden" name="selDId" value="<%=session("ssBctId")%>">
				&nbsp;<strong><%=oMainContents.FOneItem.Flastupdate%></strong>
			<% Else %>
				<input type="hidden" name="selDId" value="">
			<% End If %>
		</td>
	</tr>

	<tr bgcolor="#FFFFFF">
		<td width="150" bgcolor="#DDDDFF">��뿩��</td>
		<td>
			<% if oMainContents.FOneItem.Fisusing="N" then %>
				<input type="radio" name="isusing" value="Y">�����
				<input type="radio" name="isusing" value="N" checked >������
			<% else %>
				<input type="radio" name="isusing" value="Y" checked >�����
				<input type="radio" name="isusing" value="N">������
			<% end if %>
		</td>
	</tr>

	<tr bgcolor="#FFFFFF">
		<td colspan="2" align="center"><input type="button" value=" �� �� " onClick="SaveMainContents(frmcontents);"></td>
	</tr>
</form>
</table>
<%
	set oposcode = Nothing
	set oMainContents = Nothing
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
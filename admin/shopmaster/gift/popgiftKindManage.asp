<%@ language=vbscript %>
<% option explicit
	Response.Expires = -1440
	Response.CacheControl = "no-cache"
	Response.AddHeader "Pragma", "no-cache"
%>
<%
'####################################################
' Description :  ����ǰ ���� ���
' History : 2008.04.02 ������ ����
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/util/datelib.asp"-->
<!-- #include virtual="/lib/classes/items/itemgiftcls.asp"-->


<%
 Dim clsGift, clsGiftOpt, sViewMode, sMode
 Dim strTxt,strImg,iitemid,igkCode, S120
 Dim S401, S402, S403, S404, S405
 Dim arrList, intLoop, addimageCnt
 addimageCnt = 0
 igkCode = requestCheckVar(Request("iGK"),10)

 Dim k


 	set clsGift = new CGift
 		sMode = "KM"
 		clsGift.FGKindCode = igkCode
 		clsGift.fnGetGiftKindConts
 		strTxt = clsGift.FGKindName
		strImg = clsGift.FGKindImg
		iitemid= clsGift.FItemid
		S120 = clsGift.Fimage120

		clsGift.fnGetGiftKindAddImage
		addimageCnt = clsGift.FResultCount

		for k=0 to addimageCnt-1
		    if k=0 then S401=clsGift.Fimage400List(k)
		    if k=1 then S402=clsGift.Fimage400List(k)
		    if k=2 then S403=clsGift.Fimage400List(k)
		    if k=3 then S404=clsGift.Fimage400List(k)
		    if k=4 then S405=clsGift.Fimage400List(k)
	    next

 	set clsGift = nothing

    set clsGiftOpt = new CGift
    clsGiftOpt.FGKindCode = igkCode
    clsGiftOpt.fnGetGiftKindOptions

 Dim eFolder : eFolder =   igkCode
 Dim i, lastopt

lastopt = 0
for i=0 to clsGiftOpt.FResultCount-1
	if (clsGiftOpt.FItemList(i).Fgift_kind_option*1 > lastopt*1) then
		lastopt = clsGiftOpt.FItemList(i).Fgift_kind_option
	end if
next

%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script language="javascript">
<!--

function fnSubmit(){
	var frm = document.frmGift;


	if (frm.gift_kind_option){
		if (frm.gift_kind_option.length){
			for(var i=0;i<frm.gift_kind_option.length;i++){
				if (frm.gift_kind_optionName[i].value.length<1){
					alert('�ɼǸ��� �Է��ϼ���.');
					frm.gift_kind_optionName[i].focus();
					return;
				}
			}
		}else{
			if (frm.gift_kind_optionName.value.length<1){
				alert('�ɼǸ��� �Է��ϼ���.');
				frm.gift_kind_optionName.focus();
				return;
			}
		}
	}else{

	}

	if(confirm('���� �Ͻðڽ��ϱ�?.')){
		frm.submit();
	}
}

// �˻�
function jsSearch(){
	if(!document.frmSearch.sGKN.value){
		alert("����ǰ�������� �Է����ּ���");
		return;
	}

	document.frmSearch.submit();
}


// ��� �Ǵ� �˻� ȭ������ ����
function jsChangeMode(sViewMode){
	if (sViewMode ==""){
	document.frmSearch.sGKN.value="";
	}
	document.frmSearch.sVM.value = sViewMode;
	document.frmSearch.submit();
}

// ����ǰ �������
function jsSubmitGiftKind(){
	var frm = document.frmGift;
	if(!frm.sGKN.value){
		 alert("����ǰ�������� �Է����ּ���");
		 frm.sGKN.focus();
		 return false;
	}

	return;
}

//�˻��� ����ǰ���� ����
function jsSetGiftKind(igk, skn,strImg){
	opener.document.all.iGK.value = igk;
	opener.document.all.sGKN.value= skn;
	if(strImg !=""){
	opener.document.all.spanImg.innerHTML = "<a href=javascript:jsImgView('"+strImg+"')><img src='"+strImg+"' border=0></a>";
	}
	window.close();
}

//-- jsImgView : �̹��� Ȯ��ȭ�� ��â���� �����ֱ� --//
function jsImgView(sImgUrl){
 var wImgView;
 wImgView = window.open('/lib/showimage.asp?img='+sImgUrl,'pImg','width=100,height=100');
 wImgView.focus();
}


function jsSetImg(){
	document.domain ="10x10.co.kr";
	var winImg;
	winImg = window.open('popgiftkindupload.asp','popImg','width=370,height=150');
	winImg.focus();
}


function fnAddImage(strImg){
	document.domain ="10x10.co.kr";
	document.frmGift.sGKImg.value = strImg;
	document.all.spanImg.innerHTML = "<img src='"+strImg+"' border=0 width='60' height='30'>";
}

function fnAddImage2(strImg,sName,sSpan){
	document.domain ="10x10.co.kr";
	eval("document.frmGift." + sName).value = strImg;
	eval("document.all." + sSpan ).innerHTML = "<img src='"+strImg+"' border=0 width='60' height='30'>";
}

function jsSetImg2(sFolder, sImg, sName, sSpan){
	document.domain ="10x10.co.kr";
	var winImg;

	winImg = window.open('popgiftkindupload.asp?sF='+sFolder+'&sImg='+sImg+'&sName='+sName+'&sSpan='+sSpan,'popImg','width=370,height=150');
	winImg.focus();

	//winImg = window.open('/admin/eventmanage/common/pop_event_uploadimg.asp?sF='+sFolder+'&sImg='+sImg+'&sName='+sName+'&sSpan='+sSpan,'popImg','width=370,height=150');
	//winImg.focus();
}

function jsDelImg(sName, sSpan){
	if(confirm("�̹����� �����Ͻðڽ��ϱ�?\n\n���� �� �����ư�� ������ ó���Ϸ�˴ϴ�.")){
	   eval("document.all."+sName).value = "";
	   eval("document.all."+sSpan).style.display = "none";
	}
}

var opCd = "";
function addoption(tgt, lastopt){
	if (opCd=="") opCd = lastopt;

	if (opCd==""){
		opCd="0001";
	}else{
		opCd = (opCd*1+1);

		if (opCd<10){
			opCd = "000" + opCd;
		}else if(opCd<100){
			opCd = "00" + opCd;
		}else if(opCd<1000){
			opCd = "0" + opCd;
		}else{

		}

	}

	var tfrm = $("#"+tgt);
	var sIsrt = "<div>"
		sIsrt+= "<input class='input_a' type='text' maxlength='4' size='4' name='gift_kind_option' value='" + opCd + "' ReadOnly > "
		sIsrt+= "<input class='input_a' type='text' maxlength='32' size='20' name='gift_kind_optionName' value=''> "
		sIsrt+= "<input class='input_a' type='text' maxlength='4' size='4' name='gift_kind_Limit' value='0'> - "
		sIsrt+= "<input class='input_a' type='text' maxlength='4' size='4' name='gift_kind_LimitSold' value='0'> "
		sIsrt+= "<input type='radio' name='gift_kind_optionUsing_" + opCd + "' value='Y' checked>��� "
		sIsrt+= "<input type='radio' name='gift_kind_optionUsing_" + opCd + "' value='N'>�̻�� "
		sIsrt+= "<input class='text_ro' type='text' maxlength='2' size='2' name='prd_itemgubun' value='' ReadOnly > "
		sIsrt+= "<input class='text_ro' type='text' maxlength='8' size='8' name='prd_itemid' value='' ReadOnly > "
		sIsrt+= "<input class='text_ro' type='text' maxlength='4' size='4' name='prd_itemoption' value='' ReadOnly > "
		sIsrt+= "<input type='button' class='button' value='�˻�' onClick='jsPopSearchGiftItem(\"" + opCd + "\")' >"
		sIsrt+= "<input type='hidden' name='gift_kind_LimitYN' value='Y'>"
		sIsrt+= "</div>";
	tfrm.append(sIsrt);
}

var currentSearchItemOption = "";
function jsPopSearchGiftItem(itemoption) {
	var pop;

	currentSearchItemOption = itemoption;

	winImg = window.open("/admin/shopmaster/gift/popgiftitemlist.asp?itemgubun=85",'jsPopSearchGiftItem','width=1000,height=600,scrollbars=yes');
	winImg.focus();
}

function ReActWithThis(itemgubun, itemid, itemoption) {
	var frm = document.frmGift;

	if (frm.gift_kind_option){
		if (frm.gift_kind_option.length) {
			for(var i = 0; i < frm.gift_kind_option.length; i++) {
				if (frm.gift_kind_option[i].value == currentSearchItemOption) {
					frm.prd_itemgubun[i].value = itemgubun;
					frm.prd_itemid[i].value = itemid;
					frm.prd_itemoption[i].value = itemoption;
					return;
				}
			}
		}else{
			if (frm.gift_kind_option.value == currentSearchItemOption) {
				frm.prd_itemgubun.value = itemgubun;
				frm.prd_itemid.value = itemid;
				frm.prd_itemoption.value = itemoption;
				return;
			}
		}
	}else{

	}
}

//-->
</script>
<div style="padding: 0 5 5 5"> <img src="/images/icon_arrow_link.gif" align="absmiddle"> ����ǰ���� ���� - ��ü�̺�Ʈ��</div>
<table width="100%" border="0" align="left" class="a" cellpadding="3" cellspacing="0" >

<tr>
	<td>
		<form name="frmGift" method="post" action="giftProc.asp" onSubmit="return jsSubmitGiftKind();" style="margin:0px;">
		<input type="hidden" name="sM" value="<%=sMode%>">
		<input type="hidden" name="sGKImg" value="<%=strImg%>">
		<input type="hidden" name="iGK" value="<%=igkCode%>">
		<input type="hidden" name="S120" value="<%=S120%>">
		<input type="hidden" name="S401" value="<%=S401%>">
		<input type="hidden" name="S402" value="<%=S402%>">
		<input type="hidden" name="S403" value="<%=S403%>">
		<input type="hidden" name="S404" value="<%=S404%>">
		<input type="hidden" name="S405" value="<%=S405%>">
		<table width="100%" border="0" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
		<tr>
			<td align="center" width="100" bgcolor="<%= adminColor("tabletop") %>">����ǰ�ڵ�</td>
			<td bgcolor="#FFFFFF"><%=igkCode%></td>
		</tr>
		<tr>
			<td align="center" width="100" bgcolor="<%= adminColor("tabletop") %>">����ǰ������</td>
			<td bgcolor="#FFFFFF"><input type="text" name="sGKN" size="40" maxlength="60" value="<%=strTxt%>"></td>
		</tr>
		<tr>
			<td align="center" bgcolor="<%= adminColor("tabletop") %>">itemid</td>
			<td bgcolor="#FFFFFF"><input type="text" name="itemid" size="10" value="<%=iitemid%>"></td>
		</tr>
		<tr>
			<td align="center" bgcolor="<%= adminColor("tabletop") %>">�̹���<br>(�̺�Ʈ�� ����ǰ)</td>
			<td bgcolor="#FFFFFF">
			    <input type="button" class="button" value="�̹������" onClick="jsSetImg2('<%=eFolder%>','<%=strImg%>','sGKImg','spanImg');" >
			    <div id="spanImg">
			    <%IF strImg <> "" THEN%>
			    <a href="javascript:jsImgView('<%=strImg%>');"><img src="<%=strImg%>" width="60" height="30" border="0"></a>
			    <a href="javascript:jsDelImg('sGKImg','spanImg');"><img src="/images/icon_delete2.gif" border="0"></a>
			    <%END IF%>
			    </div>

		    </td>
		</tr>
		<tr>
		    <td align="center" bgcolor="#55AA55">�ɼǰ���</td>
		    <td bgcolor="#FFFFFF">
		        <table width="100%" border="0" cellspacing="2" cellpadding="1" class="a">
		        <tr>
		            <td width="50">�ɼ��ڵ�</td>
		            <td width="70">�ɼǸ�</td>
		            <td width="60">����</td>
		            <td width="70">��뿩��</td>
					<td width="120">�����ڵ�</td>
		            <td>
						<input type="button" class="button" value="�ɼ��߰�" onclick="javascript:addoption('optlist','<%= lastopt %>');">
					</td>
		        </tr>
		        <tr>
		            <td colspan="5">
		                <div class="a" id="optlist">
		                <% for i=0 to clsGiftOpt.FResultCount-1 %>
						<div>
						<input class="input_a" type="text" maxlength="4" size="4" name="gift_kind_option" value="<%= clsGiftOpt.FItemList(i).Fgift_kind_option %>" ReadOnly >
						<input class="input_a" type="text" maxlength="32" size="20" name="gift_kind_optionName" value="<%= clsGiftOpt.FItemList(i).Fgift_kind_optionName %>">
						<input class="input_a" type="text" maxlength="4" size="4" name="gift_kind_Limit" value="<%= clsGiftOpt.FItemList(i).Fgift_kind_Limit %>"> -
						<input class="input_a" type="text" maxlength="4" size="4" name="gift_kind_LimitSold" value="<%= clsGiftOpt.FItemList(i).Fgift_kind_LimitSold %>">
                        <input type="radio" name="gift_kind_optionUsing_<%= clsGiftOpt.FItemList(i).Fgift_kind_option %>" value="Y" <%= CHKIIF(clsGiftOpt.FItemList(i).Fgift_kind_optionUsing="Y","checked","") %>>���
                        <input type="radio" name="gift_kind_optionUsing_<%= clsGiftOpt.FItemList(i).Fgift_kind_option %>" value="N" <%= CHKIIF(clsGiftOpt.FItemList(i).Fgift_kind_optionUsing="N","checked","") %>>�̻��
						<input type="text" class="text_ro" name="prd_itemgubun" size="2" value="<%= clsGiftOpt.FItemList(i).Fprd_itemgubun %>" readonly>
						<input type="text" class="text_ro" name="prd_itemid" size="8" value="<%= clsGiftOpt.FItemList(i).Fprd_itemid %>" readonly>
						<input type="text" class="text_ro" name="prd_itemoption" size="4" value="<%= clsGiftOpt.FItemList(i).Fprd_itemoption %>" readonly>
						<input type="button" class="button" value="�˻�" onClick="jsPopSearchGiftItem('<%= clsGiftOpt.FItemList(i).Fgift_kind_option %>');" >
                        <input type="hidden" name="gift_kind_LimitYN" value="<%= clsGiftOpt.FItemList(i).Fgift_kind_LimitYN %>">
						</div>
						<% lastopt = clsGiftOpt.FItemList(i).Fgift_kind_option %>
						<% next %>
					    </div>
		            </td>
		            <td></td>
		        </tr>
		    	</table>
				(*�����ڵ�� �������� ����ǰ�� ����ϴ� ��� �Է��ϼ���)

		    </td>
		</tr>
		<tr>
			<td align="center" bgcolor="#55AAAA">��ٱ��Ͼ�����<br>(120x120)</td>
			<td bgcolor="#FFFFFF">

			        <input type="button" class="button" value="�̹������" onClick="jsSetImg2('<%=eFolder%>','<%=S120%>','S120','spanS120')">
		   		    (��ٱ��Ͽ� ǥ�õǴ� �̹���)
		   		    <div id="spanS120" style="padding: 5 5 5 5">
		   				<%IF S120 <> "" THEN %>
		   				<img  src="<%=S120%>" width="120">
		   				<a href="javascript:jsDelImg('S120','spanS120');"><img src="/images/icon_delete2.gif" border="0"></a>
		   				<%END IF%>
		   			</div>
			</td>
		</tr>

		<tr>
			<td align="center" bgcolor="#55AAAA">���˾�-1<br>(400x400)</td>
			<td bgcolor="#FFFFFF">

			        <input type="button" class="button" value="�̹������" onClick="jsSetImg2('<%=eFolder%>','<%=S401%>','S401','spanS401')">
		   		    (��ٱ��Ͽ��� �˾����� ǥ�õǴ� �̹���)
		   		    <div id="spanS401" style="padding: 5 5 5 5">
		   				<%IF S401 <> "" THEN %>
		   				<img  src="<%=S401%>" width="120">
		   				<a href="javascript:jsDelImg('S401','spanS401');"><img src="/images/icon_delete2.gif" border="0"></a>
		   				<%END IF%>
		   			</div>
			</td>
		</tr>

		<tr>
			<td align="center" bgcolor="#55AAAA">���˾�-2<br>(400x400)</td>
			<td bgcolor="#FFFFFF">

			        <input type="button" class="button" value="�̹������" onClick="jsSetImg2('<%=eFolder%>','<%=S402%>','S402','spanS402')">
		   		    (��ٱ��Ͽ��� �˾����� ǥ�õǴ� �̹���)
		   		    <div id="spanS402" style="padding: 5 5 5 5">
		   				<%IF S402 <> "" THEN %>
		   				<img  src="<%=S402%>" width="120">
		   				<a href="javascript:jsDelImg('S402','spanS402');"><img src="/images/icon_delete2.gif" border="0"></a>
		   				<%END IF%>
		   			</div>
			</td>
		</tr>

		<tr>
			<td align="center" bgcolor="#55AAAA">���˾�-3<br>(400x400)</td>
			<td bgcolor="#FFFFFF">

			        <input type="button" class="button" value="�̹������" onClick="jsSetImg2('<%=eFolder%>','<%=S403%>','S403','spanS403')">
		   		    (��ٱ��Ͽ��� �˾����� ǥ�õǴ� �̹���)
		   		    <div id="spanS403" style="padding: 5 5 5 5">
		   				<%IF S403 <> "" THEN %>
		   				<img  src="<%=S403%>" width="120">
		   				<a href="javascript:jsDelImg('S403','spanS403');"><img src="/images/icon_delete2.gif" border="0"></a>
		   				<%END IF%>
		   			</div>
			</td>
		</tr>

		<tr>
			<td align="center" bgcolor="#55AAAA">���˾�-4<br>(400x400)</td>
			<td bgcolor="#FFFFFF">

			        <input type="button" class="button" value="�̹������" onClick="jsSetImg2('<%=eFolder%>','<%=S404%>','S404','spanS404')">
		   		    (��ٱ��Ͽ��� �˾����� ǥ�õǴ� �̹���)
		   		    <div id="spanS404" style="padding: 5 5 5 5">
		   				<%IF S404 <> "" THEN %>
		   				<img  src="<%=S404%>" width="120">
		   				<a href="javascript:jsDelImg('S404','spanS404');"><img src="/images/icon_delete2.gif" border="0"></a>
		   				<%END IF%>
		   			</div>
			</td>
		</tr>

		<tr>
			<td align="center" bgcolor="#55AAAA">���˾�-5<br>(400x400)</td>
			<td bgcolor="#FFFFFF">

			        <input type="button" class="button" value="�̹������" onClick="jsSetImg2('<%=eFolder%>','<%=S405%>','S405','spanS405')">
		   		    (��ٱ��Ͽ��� �˾����� ǥ�õǴ� �̹���)
		   		    <div id="spanS405" style="padding: 5 5 5 5">
		   				<%IF S405 <> "" THEN %>
		   				<img  src="<%=S405%>" width="120">
		   				<a href="javascript:jsDelImg('S405','spanS405');"><img src="/images/icon_delete2.gif" border="0"></a>
		   				<%END IF%>
		   			</div>
			</td>
		</tr>

		<tr>
			<td colspan="2" bgcolor="#FFFFFF" align="right">
			<img src="/images/icon_confirm.gif" onClick="fnSubmit()" style="cursor:pointer">
				<!--<a href="javascript:history.back(0);"><img src="/images/icon_cancel.gif" border="0"></a>-->
			</td>
		</tr>
	</table>
	</form>
</td>
</tr>
</table>
<%
set clsGiftOpt = Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  ��ü����
' History : ���ʻ����� ��
'			2007.10.26 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/common/incSessionBctId.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/partners/partnerusercls.asp"-->
<!-- #include virtual="/lib/classes/partners/SpecialBrandCls.asp"-->
<!-- #include virtual="/lib/classes/cscenter/cs_aslistcls.asp"-->
<!-- #include virtual="/lib/util/tenEncUtil.asp" -->
<script language='javascript'>
window.resizeTo(1400,800);
</script>
<%
dim ogroup,opartner,i
dim makerid , takbae
dim groupid
dim isexposure, frequency, exposure_seq, always_exposure, startdate, enddate, brand_icon

Dim userid, encUsrId, tmpTx, tmpRn
userid = session("ssBctId")

Randomize()
tmpTx = split("A,B,C,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,W,X,Y,Z",",")
tmpRn = tmpTx(int(Rnd*26))
tmpRn = tmpRn & tmpTx(int(Rnd*26))
encUsrId = tenEnc(tmpRn & userid)

makerid = requestCheckVar(request("makerid"),32)
takbae = requestCheckVar(request("takbaebox"),32)

set opartner = new CPartnerUser
opartner.FRectDesignerID = makerid
opartner.GetOnePartnerNUser

if opartner.FResultCount < 1 then
	response.write "<script type='text/javascript'>"
	response.write "	alert('�����ϴ� �귣�尡 �ƴմϴ�. cs���� ���ο��� �� ���� ����� ������(�ѿ��)���� �ٷ� ������.');"
	response.write "</script>"
	dbget.close() : response.end
end if

set ogroup = new CPartnerGroup

if opartner.FResultCount>0 then
	ogroup.FRectGroupid = opartner.FOneItem.FGroupid
	ogroup.GetOneGroupInfo
end if


dim OReturnAddr
set OReturnAddr = new CCSReturnAddress

OReturnAddr.FRectMakerid = makerid
OReturnAddr.GetBrandReturnAddress


dim OCSBrandMemo
set OCSBrandMemo = new CCSBrandMemo

OCSBrandMemo.FRectMakerid = makerid
OCSBrandMemo.GetBrandMemo

dim brandmemo_found
if (OCSBrandMemo.Fbrandid = "") then
	brandmemo_found = "N"
else
	brandmemo_found = "Y"
end if

dim specialBrand
set specialBrand = new SpecialBrandCls
specialBrand.FRectBrandId = makerid
specialBrand.getSpecialBrandInfo()

isexposure 		= specialBrand.FOneItem.FIsexposure
frequency 		= specialBrand.FOneItem.FFrequency
always_exposure = specialBrand.FOneItem.FAlways_exposure
exposure_seq 	= specialBrand.FOneItem.FExposure_seq
startDate 		= specialBrand.FOneItem.FStartdate
endDate 		= specialBrand.FOneItem.FEnddate
brand_icon 		= specialBrand.FOneItem.FBrand_icon

if endDate <> "" then endDate = Left(specialBrand.FOneItem.FEnddate, 10)
%>
<link rel="stylesheet" href="//code.jquery.com/ui/1.12.1/themes/base/jquery-ui.css">
<link rel="stylesheet" href="/resources/demos/style.css">
<script src="https://code.jquery.com/jquery-1.12.4.js"></script>
<script src="https://code.jquery.com/ui/1.12.1/jquery-ui.js"></script>
<script type="text/javascript" src="/js/jquery.form.min.js"></script>
<script>
$( function() {
	$("#startDate").datepicker({
		dateFormat: "yy-mm-dd"
	});
	$("#endDate").datepicker({
		dateFormat: "yy-mm-dd"
	});
});
</script>
<script language="JavaScript" src="/cscenter/js/cscenter.js"></script>
<script language="JavaScript" src="/cscenter/ippbxmng/ippbxClick2Call.js"></script>
<script language="javascript">

function SaveBrandInfo(frm){
	var ret = confirm('���� �Ͻðڽ��ϱ�?');

	if(!validationChk())return false;
	if (ret){
		frm.submit();
	}
}

function jsPopCal(fName,sName)
{
	var fd = eval("document."+fName+"."+sName);

	if(fd.readOnly==false)
	{
		var winCal;
		winCal = window.open('/lib/common_cal.asp?FN='+fName+'&DN='+sName,'pCal','width=250, height=200');
		winCal.focus();
	}
}

function resizeTextArea(textarea, textareawidth) {
	var lines = textarea.value.split('\n');

	var textareaheight = 1;
	for (x = 0; x < lines.length; x++) {
		c = lines[x].length;
		if (c >= textareawidth) {
			textareaheight += Math.ceil(c / textareawidth);
		}
	}
	textareaheight += lines.length;

	if (textareaheight < 10) {
		textareaheight = 10;
	} else {
		textareaheight += 1;
	}

	textarea.rows = textareaheight;
}

function popSimpleModifyBrandInfo(makerid, mode) {
	var popwin = window.open("popsimpleModifyBrandInfo.asp?makerid=" + makerid + "&mode=" + mode,"popSimpleModifyBrandInfo","width=600,height=250,scrollbars=yes,resizable=yes");
	popwin.focus();
}
</script>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="<%= adminColor("tabletop") %>">
		<td colspan="4">
			<b>�귣�� ����</b>
		</td>
	</tr>

	<tr>
		<td colspan="4" bgcolor="#FFFFFF" height="25">[ �귣�� �⺻���� ] (������ ��ü�� �귣�庰�� ��ǰ������ �ٸ� �� �ֽ��ϴ�.)</td>
	</tr>

	<tr height="25">
		<td width="18%" bgcolor="<%= adminColor("tabletop") %>" >�귣��ID</td>
		<td width="40%" bgcolor="#FFFFFF">
            <form name=frm method="get">
                <input type="text" name="makerid" value="<%= opartner.FOneItem.FID %>" size="10">
                <input type="submit" value="�˻�">
            </form>
        </td>
		<td width="18%" bgcolor="<%= adminColor("tabletop") %>">��Ʈ��Ʈ��</td>
		<td bgcolor="#FFFFFF"><b><%= opartner.FOneItem.Fsocname_kor %></b></td>
	</tr>

	<tr height="5">
		<td colspan="4" bgcolor="#FFFFFF"></td>
	</tr>

	<form name=frmcall method=post action="return false;">
		<input type="hidden" name="returnPhone" value="<%= OReturnAddr.FreturnPhone %>">
		<input type="hidden" name="returnhp" value="<%= OReturnAddr.Freturnhp %>">
		<input type="hidden" name="csPhone" value="<%= OCSBrandMemo.FcsPhone %>">
		<input type="hidden" name="cshp" value="<%= OCSBrandMemo.Fcshp %>">
	<tr height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">��ǰ�����</td>
		<td bgcolor="#FFFFFF">
			<div style="float: left; line-height: 20px;"><%= OReturnAddr.FreturnName %></div>
			<div style="float: right;"><input type="button" class="button" value="����" onClick="popSimpleModifyBrandInfo('<%= makerid %>', 'modifyReturnCharge')"></div>
		</td>
		<td bgcolor="<%= adminColor("tabletop") %>">��ǰ��ȭ</td>
		<td bgcolor="#FFFFFF">
			<%= OReturnAddr.FreturnPhone %>
			&nbsp;
			<a href="javascript:fnClick2Call(frmcall.returnPhone);"><font color="red">[CALL]</font></a>
		</td>
	</tr>
	<tr height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">��ǰ�ڵ���</td>
		<td bgcolor="#FFFFFF">
			<%= OReturnAddr.Freturnhp %>
			&nbsp;
			<a href="javascript:fnClick2Call(frmcall.returnhp);"><font color="red">[CALL]</font></a>
		</td>
		<td bgcolor="<%= adminColor("tabletop") %>">��ǰ�̸���</td>
		<td bgcolor="#FFFFFF"><%= OReturnAddr.FreturnEmail %></td>
	</tr>

	<tr height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">��ǰ �ּ�</td>
		<td colspan="3" bgcolor="#FFFFFF" >
			[<%= OReturnAddr.FreturnZipcode %>] <%= OReturnAddr.FreturnZipaddr %> <%= OReturnAddr.FreturnEtcaddr %>
			<input type="button" class="button" value="SMS�߼�" onClick="PopCSSMSSendNew('', '','','<%= makerid %>','','','')">

			<% if OReturnAddr.flastInfoChgDT<>"" and not(isNull(OReturnAddr.flastInfoChgDT)) then %>
				&nbsp;&nbsp;
				���������� : <%= OReturnAddr.flastInfoChgDT %>
			<% end if %>
		</td>
	</tr>

	<tr height="5">
		<td colspan="4" bgcolor="#FFFFFF"></td>
	</tr>

	<tr height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">CS�����</td>
		<td bgcolor="#FFFFCC">
			<div style="float: left; line-height: 20px;"><%= OCSBrandMemo.FcsName %></div>
			<div style="float: right;"><input type="button" class="button" value="����" onClick="popSimpleModifyBrandInfo('<%= makerid %>', 'modifyCSCharge')"></div>
		</td>
		<td bgcolor="<%= adminColor("tabletop") %>">CS��ȭ</td>
		<td bgcolor="#FFFFCC">
			<%= OCSBrandMemo.FcsPhone %>
			&nbsp;
			<a href="javascript:fnClick2Call(frmcall.csPhone);"><font color="red">[CALL]</font></a>
		</td>
	</tr>
	<tr height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">CS�ڵ���</td>
		<td bgcolor="#FFFFCC">
			<%= OCSBrandMemo.Fcshp %>
			&nbsp;
			<a href="javascript:fnClick2Call(frmcall.cshp);"><font color="red">[CALL]</font></a>
		</td>
		<td bgcolor="<%= adminColor("tabletop") %>">CS�̸���</td>
		<td bgcolor="#FFFFCC"><%= OCSBrandMemo.FcsEmail %></td>
	</tr>
	</form>

	<tr height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">��������</td>
		<td bgcolor="#FFFFCC"><%= OCSBrandMemo.FcsModifyDay %></td>
		<td bgcolor="<%= adminColor("tabletop") %>"></td>
		<td bgcolor="#FFFFCC"><%= OCSBrandMemo.FcsReguserID %></td>
	</tr>

	<tr height="25">
		<td colspan="4" bgcolor="#FFFFFF" height="25">[ �귣�� ������� ]</td>
	</tr>
	<tr height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">���ǹ�ۿ���</td>
		<td bgcolor="#FFFFFF">
			<% if (opartner.FOneItem.IsFreeBeasong) then %>
				�׻� ������
			<% end if %>
			<% if (opartner.FOneItem.IsUpcheReceivePayDeliverItem) then %>
				���ҹ��
			<% end if %>
			<% if opartner.FOneItem.IsUpcheParticleDeliverItem then %>
				���ݺ� ������
			<% end if %>
			<% if ((opartner.FOneItem.IsUpcheParticleDeliverItem) or (opartner.FOneItem.IsUpcheReceivePayDeliverItem)) and Not(opartner.FOneItem.IsFreeBeasong) then %>
			<% else %>
				N
			<% end if %>
		</td>
		<td bgcolor="<%= adminColor("tabletop") %>">��ۺ�</td>
		<td bgcolor="#FFFFFF">
			<% if opartner.FOneItem.IsUpcheParticleDeliverItem then %>
			<b><%=FormatNumber(opartner.FOneItem.FdefaultFreeBeasongLimit,0)%></b>�� �̸� <b><%=FormatNumber(opartner.FOneItem.FdefaultDeliverPay,0)%></b> ��
			<% else %>
			<% if IsNull(opartner.FOneItem.FdefaultDeliverPay) then opartner.FOneItem.FdefaultDeliverPay = 0 end if %>
			��ǰ��ۺ� <b><%=FormatNumber(opartner.FOneItem.FdefaultDeliverPay,0)%></b> ��
			<% end if %>
		</td>
	</tr>
	<tr height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">�ŷ��ù��</td>
		<td bgcolor="#FFFFFF" colspan="3">
			<%= opartner.FOneItem.Ftakbae_name %> (<%= opartner.FOneItem.Ftakbae_tel %>)
			<%
			Select Case OCSBrandMemo.Fis_return_allow
				Case "Y"
			%>
			<font color="blue"><b>��ü����ȸ��</b></font>
			<%
				Case "N"
			%>
			<font color="red"><b>��üȸ���Ұ�</b></font>
			<%
				Case Else
					''//
			End Select
			%>
		</td>
	</tr>
	<form name="brandmemo" method="post" action="do_brandmemo_input.asp">
	<tr height="25">
		<td colspan="4" bgcolor="#FFFFFF" height="25">[ �귣�� �߰����� ]</td>
	</tr>
	<tr height="25">
		<input type=hidden name=makerid value="<%= makerid %>">
		<input type=hidden name='isSpecialBrand' value="<%= specialbrand.FResultCount %>">
		<input type=hidden name=mode value="<% if brandmemo_found = "Y" then %>modify<% else %>insert<% end if %>">

		<td bgcolor="<%= adminColor("tabletop") %>">ī�װ�</td>
		<td bgcolor="#FFFFCC"><% SelectBoxBrandCategory "catecode", opartner.FOneItem.Fcatecode %></td>
		<td bgcolor="<%= adminColor("tabletop") %>">���MD</td>
		<td bgcolor="#FFFFCC"><% drawSelectBoxCoWorker_OnOff "mduserid", opartner.FOneItem.Fmduserid, "on" %></td>
	</tr>

	<tr height="25">
		<td bgcolor="<%= adminColor("tabletop") %>"><b>ȸ������</b></td>
		<td bgcolor="#FFFFCC" colspan="3">
			<select class="select" name="is_return_allow">
		     	<option value="-" >-</option>
		     	<option value="Y" <% if (OCSBrandMemo.Fis_return_allow = "Y") then %>selected<% end if %>>����</option>
		     	<option value="N" <% if (OCSBrandMemo.Fis_return_allow = "N") then %>selected<% end if %>>�Ұ���</option>
	     	</select>
			<input type="text" size="40" name="ret_comment" value="<%= OCSBrandMemo.Freturn_comment %>">
	    </td>
	</tr>

	<tr height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">��㰡�ɽð�</td>
		<td bgcolor="#FFFFFF" colspan="3">
			<select class="select" name="tel_start">
				<option value="0">-- : --</option>
		     	<% for i = 6 to 15 %>
		     	<option value="<%= i %>" <% if (OCSBrandMemo.Ftel_start = i) then %>selected<% end if %>><%= i %>:00</option>
		    	<% next %>
	     	</select>
	     	~
			<select class="select" name="tel_end">
				<option value="0">-- : --</option>
		     	<% for i = 12 to 21 %>
		     	<option value="<%= i %>" <% if (OCSBrandMemo.Ftel_end = i) then %>selected<% end if %>><%= i %>:00</option>
		    	<% next %>
	     	</select>
	      	(����� �ٹ�����
	      	<select class="select" name="is_saturday_work">
		     	<option value="-" >-</option>
		     	<option value="Y" <% if (OCSBrandMemo.Fis_saturday_work = "Y") then %>selected<% end if %>>Y</option>
		     	<option value="N" <% if (OCSBrandMemo.Fis_saturday_work = "N") then %>selected<% end if %>>N</option>
	     	</select>)
	     </td>
	</tr>

	<tr height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">���ɽð�</td>
		<td bgcolor="#FFFFFF" colspan="3">
			<select class="select" name="lunch_start">
				<option value="0">-- : --</option>
		     	<% for i = 6 to 15 %>
		     	<option value="<%= i %>" <% if (OCSBrandMemo.Flunch_start = i) then %>selected<% end if %>><%= i %>:00</option>
		    	<% next %>
	     	</select>
	     	~
			<select class="select" name="lunch_end">
				<option value="0">-- : --</option>
		     	<% for i = 12 to 21 %>
		     	<option value="<%= i %>" <% if (OCSBrandMemo.Flunch_end = i) then %>selected<% end if %>><%= i %>:00</option>
		    	<% next %>
	     	</select>
	     </td>
	</tr>

	<tr height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">�ް�����</td>
		<td bgcolor="#FFFFFF" colspan="3">
			<input type="text" size="10" name="vacation_startday" value="<%= OCSBrandMemo.Fvacation_startday %>" onClick="jsPopCal('brandmemo','vacation_startday');" style="cursor:hand;"> - <input type="text" size="10" name="vacation_endday" value="<%= OCSBrandMemo.Fvacation_endday %>" onClick="jsPopCal('brandmemo','vacation_endday');" style="cursor:hand;">
			<select class="select" name="vacation_div">
				<option value="">�ް�����</option>
		     	<option value="����" <% if (OCSBrandMemo.Fvacation_div = "����") then %>selected<% end if %> >����</option>
				<option value="����" <% if (OCSBrandMemo.Fvacation_div = "����") then %>selected<% end if %> >����</option>
				<option value="�߼�" <% if (OCSBrandMemo.Fvacation_div = "�߼�") then %>selected<% end if %> >�߼�</option>
				<option value="�ϰ��ް�" <% if (OCSBrandMemo.Fvacation_div = "�ϰ��ް�") then %>selected<% end if %> >�ϰ��ް�</option>
				<option value="��Ÿ" <% if (OCSBrandMemo.Fvacation_div = "��Ÿ") then %>selected<% end if %> >��Ÿ</option>
	     	</select>
	     </td>
	</tr>
	<tr height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">����ǰ<br />�Ұ�����</td>
		<td bgcolor="#FFFFFF" colspan="3">
			<select class="select" name="customer_return_deny">
				<option value=""></option>
                <option value="N" <% if (OCSBrandMemo.Fcustomer_return_deny = "N") then %>selected<% end if %> >��ǰ���</option>
		     	<option value="Y" <% if (OCSBrandMemo.Fcustomer_return_deny = "Y") then %>selected<% end if %> >�� ������ǰ �Ұ�</option>
	     	</select>
	     </td>
	</tr>
	<tr height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">��Ÿ�޸�</td>
		<td bgcolor="#FFFFFF" colspan="3">
		<textarea class="textarea" name=brand_comment cols="70" rows="10"><% if (OCSBrandMemo.Fbrand_comment = "") then %>�����޸�(��󿬶���,ȯ�Ұ���,�±�ȯ���ɿ��� ��)<% else %><%= OCSBrandMemo.Fbrand_comment %><% end if %></textarea>
		</td>
	</tr>
	<!-- special brand ����� �귣�� 2019-07-15 -->
	<!--======================================================================================================-->
	<tr height="25" bgcolor="<%= adminColor("tabletop") %>">
		<td colspan="4">
			<b>����� �귣�� ����</b>
		</td>
	</tr>
	<script>
	$(function(){
		if('<%=isexposure%>' == '1'){
			showSpecialBrandInfo(true)
		}else{
			showSpecialBrandInfo(false)
		}
		if('<%=always_exposure%>' == '1'){
			showPeriod(true)
		}else{
			showPeriod(false)
		}


	})
	function showSpecialBrandInfo(isExposed){
		if(isExposed){
			$(".expose").css("display", "")
			$("#selectExposure input[value=1]").attr('checked', true)
		}else{
			$(".expose").css("display", "none")
			$("#selectExposure input[value=0]").attr('checked', true)
		}
	}

	function showPeriod(isExposed){
		if(isExposed){
			$("#calendar").css("display", "none")
			$("#selectPeriod input[value=1]").attr('checked', true)
		}else{
			$("#calendar").css("display", "")
			$("#selectPeriod input[value=0]").attr('checked', true)
		}
	}

	function validationChk(){
		var tmpFrm = document.brandmemo
		if(document.brandmemo.isexposure[1].checked) return true;
		if(tmpFrm.brand_icon.value == ""){
			alert("������ �̹����� �־��ּ���.")
			return false;
		}
		if(tmpFrm.frequency.value == ""){
			alert("���� �󵵸� �������ּ���.")
			return false;
		}
		if(tmpFrm.exposure_seq.value == ""){
			alert("���� ������ �����ּ���.")
			return false;
		}
		if(tmpFrm.always_exposure.value == ""){
			alert("��� ���� ���θ� �������ּ���.")
			return false;
		}
		if(tmpFrm.always_exposure.value == 0){
			if(tmpFrm.startDate.value == ""){
				alert("���� �������� �־��ּ���.")
				return false;
			}
			if(tmpFrm.endDate.value == ""){
				alert("���� �������� �־��ּ���.")
				return false;
			}
		}

		return true
	}

	function jsCheckUpload() {
		var gubun = document.frmUpload.imgtype.value;
		var test = $("input[id="+gubun+"]").val();
		console.log(gubun);
		console.log(test);
		// return false;
		if($("#fileupload").val()!="") {
			$("#fileupmode").val("upload");

			$('#ajaxform').ajaxSubmit({
				//�������� validation check�� �ʿ��Ұ��
				beforeSubmit: function (data, frm, opt) {
					if(!(/\.(jpg|jpeg|png)$/i).test(frm[0].upfile.value)) {
						alert("JPG,PNG �̹������ϸ� ���ε� �Ͻ� �� �ֽ��ϴ�.");
						$("#fileupload").val("");
						return false;
					}
				},
				//submit������ ó��
				success: function(responseText, statusText){
					var resultObj = JSON.parse(responseText)

					if(resultObj.response=="fail") {
						alert(resultObj.faildesc);
					} else if(resultObj.response=="ok") {
						$("#filepre").val(resultObj.fileurl);
						$("img[id="+gubun+"src]").hide().attr("src",$("#filepre").val()+"?"+Math.floor(Math.random()*1000)).fadeIn("fast");
						$("input[id="+gubun+"]").val(resultObj.fileurl);
					} else {
						alert("ó���� ������ �߻��߽��ϴ�.\n" + responseText);
					}
					$("#fileupload").val("");
				},
				//ajax error
				error: function(err){
					alert("ERR: " + err.responseText);
					$("#fileupload").val("");
				}
			});
		}
	}
	function setImgType(type){
		console.log(document.frmUpload)
		document.frmUpload.imgtype.value = type;
		return false;
	}
	</script>
	<tr height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">���� ����</td>
		<td bgcolor="#FFFFFF" colspan="3" id="selectExposure">
			<input type="radio" onclick="showSpecialBrandInfo(true)" name="isexposure" value="1" class="radio" />������&nbsp;&nbsp;
			<input type="radio" onclick="showSpecialBrandInfo(false)" name="isexposure" value="0" class="radio" />�������&nbsp;&nbsp;
	    </td>
	</tr>
	<tr height="25" class="expose">
		<td bgcolor="<%= adminColor("tabletop") %>">�귣�� ������ �̹���</td>
		<td bgcolor="#FFFFFF" colspan="3" id="selectExposure">
					<div class="inTbSet">
						<div id="type1Img1">
							<p class="registImg">
								<input type="hidden" id="brand_icon" name="brand_icon" value="<%=brand_icon%>" />
								<img id="brand_iconsrc" src="<%=chkIIF(brand_icon="" or isNull(brand_icon),"/images/admin_login_logo2.png",brand_icon)%>" style="height:118px; border:1px solid #EEE;"/>
								<button onclick="setImgType('brand_icon')" type="button">
									<label for="fileupload"><%=chkIIF(brand_icon="","�̹��� ���ε�","�̹��� ����")%></label>
								</button>
							</p>
						</div>
					</div>
	    </td>
	</tr>
	<tr height="25" class="expose">
		<td bgcolor="<%= adminColor("tabletop") %>">���� ��</td>
		<td bgcolor="#FFFFFF" colspan="3">
			<input type="radio" name="frequency" value="60" class="radio" <%=chkIIF(frequency="60", "checked", "")%>/>����&nbsp;&nbsp;
			<input type="radio" name="frequency" value="30" class="radio" <%=chkIIF(frequency="30", "checked", "")%>/>�⺻&nbsp;&nbsp;
			<input type="radio" name="frequency" value="10" class="radio" <%=chkIIF(frequency="10" or frequency="", "checked", "")%>/>����&nbsp;&nbsp;
	    </td>
	</tr>
	<tr height="25" class="expose">
		<td bgcolor="<%= adminColor("tabletop") %>">���� ����</td>
		<td bgcolor="#FFFFFF" colspan="3">
			<input type="number" name="exposure_seq" value="<%=chkIIF(exposure_seq<>"",exposure_seq ,0)%>" class="radio" />
	    </td>
	</tr>
	<tr height="25" class="expose">
		<td bgcolor="<%= adminColor("tabletop") %>">���� �Ⱓ</td>
		<td bgcolor="#FFFFFF" colspan="3" id="selectPeriod">
			<input type="radio" onclick="showPeriod(true)" name="always_exposure" value="1" class="radio" <%=chkIIF(always_exposure="1" or always_exposure="", "checked", "")%>/>��� ����<br>
			<input type="radio" onclick="showPeriod(false)" name="always_exposure" value="0" class="radio" <%=chkIIF(always_exposure="0", "checked", "")%>/>���� �Ⱓ ����<br>
			<!-- �޷� -->
			<div id="calendar">
				<input type="text" id="startDate" name="startDate" value="<%=startDate%>" style="width:80px" readOnly>
				-
				<input type="text" id="endDate" name="endDate" value="<%=endDate%>" style="width:80px" readOnly>
			</div>
			<!-- �޷� -->
	    </td>
	</tr>

	<!--======================================================================================================-->
	<tr height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">����������</td>
		<td bgcolor="#FFFFFF" colspan="3">
		<% if Len(OCSBrandMemo.Flast_modifyday) > 10 then %>
			<%= Left(OCSBrandMemo.Flast_modifyday) %>
		<% else %>
			<%= (OCSBrandMemo.Flast_modifyday) %>
		<% end if %>
		</td>
	</tr>
	<tr height="30" align="center">
		<td colspan="4" bgcolor="#FFFFFF" height="25">
			<input type="button" class="button" value="�߰���������" onclick="SaveBrandInfo(brandmemo)"></td>
		</td>
	</tr>
	</form>


	<tr height="25">
		<td colspan="4" bgcolor="#FFFFFF" height="25">[ ��ü�⺻���� ]</td>
	</tr>
	<tr height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">ȸ���(��ȣ)</td>
		<td bgcolor="#FFFFFF"><b><%= ogroup.FOneItem.FCompany_name %></b></td>
		<td bgcolor="<%= adminColor("tabletop") %>">�׷��ڵ�</td>
		<td bgcolor="#FFFFFF"><b><%= opartner.FOneItem.FGroupid %></b></td>
	</tr>
	<tr height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">��ǥ��ȭ</td>
		<td bgcolor="#FFFFFF"><%= ogroup.FOneItem.Fcompany_tel %></td>
		<td bgcolor="<%= adminColor("tabletop") %>">�ѽ�</td>
		<td bgcolor="#FFFFFF"><%= ogroup.FOneItem.Fcompany_fax %></td>
	</tr height="25">
	<tr height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">�繫�� �ּ�</td>
		<td bgcolor="#FFFFFF" colspan=3>[<%= ogroup.FOneItem.Freturn_zipcode %>] <%= ogroup.FOneItem.Freturn_address %> <%= ogroup.FOneItem.Freturn_address2 %></td>
	</tr height="25">



	<tr>
		<td colspan="4" bgcolor="#FFFFFF" height="25">[ ��ü ��������� ]</td>
	</tr>

	<tr height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">����ڸ�</td>
		<td bgcolor="#FFFFFF"><%= ogroup.FOneItem.Fmanager_name %></td>
		<td bgcolor="<%= adminColor("tabletop") %>">�Ϲ���ȭ</td>
		<td bgcolor="#FFFFFF"><%= ogroup.FOneItem.Fmanager_phone %></td>
	</tr>
	<tr height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">E-Mail</td>
		<td bgcolor="#FFFFFF"><%= ogroup.FOneItem.Fmanager_email %></td>
		<td bgcolor="<%= adminColor("tabletop") %>">�ڵ���</td>
		<td bgcolor="#FFFFFF"><%= ogroup.FOneItem.Fmanager_hp %></td>
	</tr>
	<tr height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">��۴���ڸ�</td>
		<td bgcolor="#FFFFFF" colspan="3">�귣�庰�� ��ȸ �����մϴ�</td>
	</tr>
	<!--
	<tr height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">��۴���ڸ�</td>
		<td bgcolor="#FFFFFF" colspan="3"><%= ogroup.FOneItem.Fdeliver_name %></td>
		<td bgcolor="<%= adminColor("tabletop") %>">�Ϲ���ȭ</td>
		<td bgcolor="#FFFFFF"><%= ogroup.FOneItem.Fdeliver_phone %></td>
	</tr>
	<tr height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">E-Mail</td>
		<td bgcolor="#FFFFFF"><%= ogroup.FOneItem.Fdeliver_email %></td>
		<td bgcolor="<%= adminColor("tabletop") %>">�ڵ���</td>
		<td bgcolor="#FFFFFF"><%= ogroup.FOneItem.Fdeliver_hp %></td>
	</tr>
	-->
	<tr height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">�������ڸ�</td>
		<td bgcolor="#FFFFFF"><%= ogroup.FOneItem.Fjungsan_name %></td>
		<td bgcolor="<%= adminColor("tabletop") %>">�Ϲ���ȭ</td>
		<td bgcolor="#FFFFFF"><%= ogroup.FOneItem.Fjungsan_phone %></td>
	</tr>
	<tr height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">E-Mail</td>
		<td bgcolor="#FFFFFF"><%= ogroup.FOneItem.Fjungsan_email %></td>
		<td bgcolor="<%= adminColor("tabletop") %>">�ڵ���</td>
		<td bgcolor="#FFFFFF"><%= ogroup.FOneItem.Fjungsan_hp %></td>
	</tr>

	<!-- CS����� ��û���� �ʺ��̰� ó��
	<tr height="25">
		<td colspan="4" bgcolor="#FFFFFF" height="25">[ ��ü ����ڵ������ ]</td>
	</tr>
	<tr height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">ȸ���(��ȣ)</td>
		<td bgcolor="#FFFFFF"><%= ogroup.FOneItem.FCompany_name %></td>
		<td bgcolor="<%= adminColor("tabletop") %>">��ǥ��</td>
		<td bgcolor="#FFFFFF"><%= ogroup.FOneItem.Fceoname %></td>
	</tr>
	<tr height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">����ڹ�ȣ</td>
		<td bgcolor="#FFFFFF" colspan="3"><%= ogroup.FOneItem.Fcompany_no %></td>
	</tr>
	<tr height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">����������</td>
		<td colspan="3" bgcolor="#FFFFFF" >
			<%= ogroup.FOneItem.Fcompany_zipcode %>&nbsp;
			<%= ogroup.FOneItem.Fcompany_address %>&nbsp;
			<%= ogroup.FOneItem.Fcompany_address2 %>
		</td>
	</tr>
	<tr height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">����</td>
		<td bgcolor="#FFFFFF"><%= ogroup.FOneItem.Fcompany_uptae %></td>
		<td bgcolor="<%= adminColor("tabletop") %>">����</td>
		<td bgcolor="#FFFFFF"><%= ogroup.FOneItem.Fcompany_upjong %></td>
	</tr>
	-->

	<tr align="center">
		<td colspan="4" bgcolor="#FFFFFF" height="30">
			<input type="button" class="button" value="�ݱ�" onclick="self.close();"></td>
		</td>
	</tr>

</table>

<%
set opartner = Nothing
set ogroup = Nothing
%>

<form name="frmUpload" id="ajaxform" action="<%=uploadImgUrl%>/linkweb/common/simpleCommonImgUploadProc.asp" method="post" enctype="multipart/form-data" style="display:none; height:0px;width:0px;">
	<input type="file" name="upfile" id="fileupload" onchange="jsCheckUpload();" accept="image/*" />
	<input type="hidden" name="mode" id="fileupmode" value="upload">
	<input type="hidden" name="div" value="TQ">
	<input type="hidden" name="upPath" value="/appmanage/specialbrand/">
	<input type="hidden" name="tuid" value="<%=encUsrId%>">
	<input type="hidden" name="prefile" id="filepre" >
	<input type="hidden" name="imgtype">
</form>
<script>

window.onload = function() {
	// resizeTextArea(document.getElementById("brand_comment"), 50);
}

</script>

<!-- #include virtual="/common/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->

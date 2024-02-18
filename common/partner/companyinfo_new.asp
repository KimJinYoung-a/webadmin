<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
Response.AddHeader "Access-Control-Allow-Origin","*"
Response.AddHeader "Access-Control-Allow-Methods","POST"
Response.AddHeader "Access-Control-Allow-Headers","X-Requested-With"
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/classes/partners/partnerusercls.asp"-->
<!-- #include virtual="/lib/classes/cscenter/cs_aslistcls.asp"-->
<!-- #include virtual="/partner/lib/adminHead.asp" -->
<!-- #include virtual="/lib/util/tenEncUtil.asp"-->
<!-- #include virtual="/lib/util/base64unicode.asp"-->
<%
dim i,page
dim pid, pcuserdiv, qs
qs = request.querystring("qs")
qs = TBTDecryptUrl(qs)
qs = split(qs,"|")

pid = requestCheckVar(qs(0),32)
pcuserdiv = requestCheckVar(qs(1),16)
if pid="" or pcuserdiv="" then
    response.write "<script>alert('�߸��� �����Դϴ�.');</script>"
    response.end
end if

dim opartner
set opartner = new CPartnerUser
	opartner.FCurrpage = 1
	opartner.FRectDesignerID = pid
	opartner.FPageSize = 1
	opartner.GetOnePartnerNUser
if opartner.FResultCount < 1 then
    response.write "<script>alert('���� ���� ��ü(�귣��)�� �ƴմϴ�.');</script>"
    response.end
end if

dim uploadImgUrl
dim C_IS_SSL_ENABLED : C_IS_SSL_ENABLED = (Request.ServerVariables("HTTPS") = "on")
IF application("Svr_Info")="Dev" THEN
	if (C_IS_SSL_ENABLED = True) then
 		uploadImgUrl 		= "https://testupload.10x10.co.kr"          '' upload.10x10.co.kr ���ؼ� Nas Server�� ���ε�
	else
 		uploadImgUrl 		= "http://testupload.10x10.co.kr"          '' upload.10x10.co.kr ���ؼ� Nas Server�� ���ε�
	end if
ELSE
	if (C_IS_SSL_ENABLED = True) then
 		uploadImgUrl 		= "https://upload.10x10.co.kr"          '' upload.10x10.co.kr ���ؼ� Nas Server�� ���ε�
	else
 		uploadImgUrl 		= "http://upload.10x10.co.kr"          '' upload.10x10.co.kr ���ؼ� Nas Server�� ���ε�
	end if
 END IF
%>
<script type="text/javascript" src="/js/jquery.form.min.js"></script>
<script type="text/javascript" scr="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript">

function SaveUpcheInfo(frm){
    
    if (frm.password.value.length<1){
        alert('�귣�� 1�� �н����带 �Է��ϼ���.');
        frm.password.focus();
        return;
	}


	if (frm.password.value.length < 8 || frm.password.value.length > 16){
        alert("�н������ ������� 8~16���Դϴ�.");
        frm.password.focus();
        return ;
	}
		 
	if (!fnChkComplexPassword(frm.password.value)) {
        alert('�н������ ����/����/Ư������ �� �� ���� �̻��� �������� �Է����ּ���.');
        frm.password.focus();
        return;
	}
			
	if(frm.password.value==frm.uid.value) {
        alert("���̵�� �ٸ� ��й�ȣ�� ������ּ���.");
        frm.password.focus();
        return  ;
	}
		 
	if (frm.passwordS.value.length<1){
        alert('�귣�� 2�� �н����带 �Է��ϼ���.');
        frm.passwordS.focus();
        return;
	}
		
	if (frm.passwordS.value.length < 8 || frm.passwordS.value.length > 16){
        alert("2�� �н������ ������� 8~16���Դϴ�.");
        frm.passwordS.focus();
        return ;
	}
		 
	if (!fnChkComplexPassword(frm.passwordS.value)) {
        alert('2�� �н������ ����/����/Ư������ �� �� ���� �̻��� �������� �Է����ּ���.');
        frm.passwordS.focus();
        return;
	}
			
	if(frm.passwordS.value==frm.uid.value) {
        alert("���̵�� �ٸ� ��й�ȣ�� ������ּ���.");
        frm.passwordS.focus();
        return  ;
	}
		 
		
	if(frm.passwordS.value==frm.password.value) {
        alert("1�� ��й�ȣ��  �ٸ� ��й�ȣ�� ������ּ���.");
        frm.passwordS.focus();
        return  ;
	}
    
    if (frm.company_name.value.length<1){
    	alert('����� ��ϻ��� ȸ����� �Է��ϼ���.');
    	frm.company_name.focus();
    	return;
    }

    if (frm.ceoname.value.length<1){
    	alert('����� ��ϻ��� ��ǥ�ڸ��� �Է��ϼ���.');
    	frm.ceoname.focus();
    	return;
    }

    if (frm.company_no.value.length<1){
    	alert('����� ��� ��ȣ�� �Է��ϼ���.');
    	frm.company_no.focus();
    	return;
    }

	//if (frm.jungsan_gubun.value.length<1){
	//	alert('���������� �����ϼ���.');
	//	frm.jungsan_gubun.focus();
	//	return;
	//}

	if (frm.company_zipcode.value.length<1){
		alert('�����ȣ�� �����ϼ���.');
		frm.company_zipcode.focus();
		return;
	}

	if (frm.company_address.value.length<1){
		alert('����� ��ϻ��� �ּ�1�� �Է��ϼ���.');
		frm.company_address.focus();
		return;
	}

	if (frm.company_address2.value.length<1){
		alert('����� ��ϻ��� �ּ�2�� �Է��ϼ���.');
		frm.company_address2.focus();
		return;
	}

	if (frm.company_uptae.value.length<1){
		alert('����� ��ϻ��� ���¸� �Է��ϼ���.');
		frm.company_uptae.focus();
		return;
	}

	if (frm.company_upjong.value.length<1){
		alert('����� ��ϻ��� ������ �Է��ϼ���.');
		frm.company_upjong.focus();
		return;
	}

    if (frm.jungsan_bank.value.length<1){
		alert('�ŷ������� �Է��ϼ���.');
		frm.jungsan_bank.focus();
		return;
	}
    
    if (frm.jungsan_acctno.value.length<1){
		alert('���¹�ȣ�� �Է��ϼ���.');
		frm.jungsan_acctno.focus();
		return;
	}
    
    if (frm.jungsan_acctname.value.length<1){
		alert('�����ָ��� �Է��ϼ���.');
		frm.jungsan_acctname.focus();
		return;
	}

    if (frm.company_tel.value.length<1){
		alert('��ǥ��ȭ ��ȣ�� �Է��ϼ���.');
		frm.company_tel.focus();
		return;
	}

	if (frm.return_zipcode.value.length<1){
		alert('�繫���ּ� �����ȣ�� �����ϼ���.');
		frm.return_zipcode.focus();
		return;
	}

	if (frm.return_address.value.length<1){
		alert('�繫���ּ�1 �� �Է��ϼ���.');
		frm.return_address.focus();
		return;
	}

	if (frm.return_address2.value.length<1){
		alert('�繫���ּ�2�� �Է��ϼ���.');
		frm.return_address2.focus();
		return;
	}

	if (frm.company_tel.value.length<1){
		alert('��Ʈ�� ��ȭ��ȣ�� �Է��ϼ���.');
		frm.company_tel.focus();
		return;
	}

	if (frm.manager_name.value.length<1){
		alert('����� ������ �Է��ϼ���.');
		frm.manager_name.focus();
		return;
	}

	if (frm.manager_phone.value.length<1){
		alert('����� ��ȭ��ȣ�� �Է��ϼ���.');
		frm.manager_phone.focus();
		return;
	}

	if (frm.manager_email.value.length<1){
		alert('����� �̸����� �Է��ϼ���.');
		frm.manager_email.focus();
		return;
	}

	if (frm.manager_hp.value.length<1){
		alert('����� �ڵ����� �Է��ϼ���.');
		frm.manager_hp.focus();
		return;
    }
    
    if (frm.defaultsongjangdiv.value.length<1){
		alert('��� �ù�縦 �������ּ���.');
		frm.defaultsongjangdiv.focus();
		return;
	}

    if (frm.company_no_img.value==""){
		alert('����� ����� �̹����� ������ּ���.');
		frm.company_no_img2.focus();
		return;
    }
    
    if (frm.jungsan_acctno_img.value==""){
		alert('���� �纻 �̹����� ������ּ���.');
		frm.jungsan_acctno_img2.focus();
		return;
	}

	var ret = confirm('��Ʈ�� ������ ���� �Ͻðڽ��ϱ�?');

	if (ret){
		frm.submit();
	}
}

function SaveBrandReturnInfo(frm){
	var ret = confirm('�귣�� ��ǰ ������ ���� �Ͻðڽ��ϱ�?');

	if (ret){
		frm.submit();
	}
}

// �н����� ���⵵ �˻�
function fnChkComplexPassword(pwd) {
    var aAlpha = /[a-z]|[A-Z]/;
    var aNumber = /[0-9]/;
    var aSpecial = /[!|@|#|$|%|^|&|*|(|)|-|_]/;
    var sRst = true;

    if(pwd.length < 8){
        sRst=false;
        return sRst;
    }

    var numAlpha = 0;
    var numNums = 0;
    var numSpecials = 0;
    for(var i=0; i<pwd.length; i++){
        if(aAlpha.test(pwd.substr(i,1)))
            numAlpha++;
        else if(aNumber.test(pwd.substr(i,1)))
            numNums++;
        else if(aSpecial.test(pwd.substr(i,1)))
            numSpecials++;
    }

    if((numAlpha>0&&numNums>0)||(numAlpha>0&&numSpecials>0)||(numNums>0&&numSpecials>0)) {
    	sRst=true;
    } else {
    	sRst=false;
    }
    return sRst;
}

function CopyZip(flag,post1,post2,add,dong){
	if (flag=="s"){
		frmupche.company_zipcode.value= post1 + "-" + post2;
		frmupche.company_address.value= add;
		frmupche.company_address2.value= dong;
	}else if(flag=="m"){
		frmupche.return_zipcode.value= post1 + "-" + post2;
		frmupche.return_address.value= add;
		frmupche.return_address2.value= dong;
	}else if(flag=="b"){
		frmbrand.return_zipcode.value= post1 + "-" + post2;
		frmbrand.return_address.value= add;
		frmbrand.return_address2.value= dong;
	}
}

function popZip(flag){
	var popwin = window.open("/lib/searchzip3.asp?target=" + flag,"searchzip3","width=460 height=240 scrollbars=yes resizable=yes");
	popwin.focus();
}
 
function SameReturnAddr(bool){
	if (bool){
		frmupche.return_zipcode.value = frmupche.company_zipcode.value;
		frmupche.return_address.value = frmupche.company_address.value;
		frmupche.return_address2.value = frmupche.company_address2.value;
	}else{
		frmupche.return_zipcode.value = "";
		frmupche.return_address.value = "";
		frmupche.return_address2.value = "";
	}
}

function SameReturnAddr2(bool){
	if (bool){
		frmupche.p_return_zipcode.value = frmupche.return_zipcode.value;
		frmupche.p_return_address.value = frmupche.return_address.value;
		frmupche.p_return_address2.value = frmupche.return_address2.value;
	}else{
		frmupche.p_return_zipcode.value = "";
		frmupche.p_return_address.value = "";
		frmupche.p_return_address2.value = "";
	}
}

//��ü ���� �����ȣ ã��
function TnFindZipPartnerDistribution(frmname, strMode){
    var TnFindZipNewPartner = window.open('/partner/lib/searchzip_company.asp?target=' + frmname + '&strMode='+strMode, 'TnFindZipNewPartner', 'width=580,height=690,left=400,top=200,scrollbars=yes,resizable=yes');
    TnFindZipNewPartner.focus();
}

function PopUpcheReturnAddrOnly(){
	var popwin = window.open("/partner/company/company_returnaddr_mod_pop.asp","popupchereturnaddronly","width=1200 height=600 scrollbars=yes resizable=yes");
	popwin.focus();
}

var orgjungsan_gubun = "�Ϲݰ���";
function fnJungsanGubunChanged() {
	var frm = document.frmupche;
	var company_no = document.getElementById("company_no");

	if ((orgjungsan_gubun != "����(�ؿ�)") && (frm.jungsan_gubun.value != "����(�ؿ�)")) {
		orgjungsan_gubun = frm.jungsan_gubun.value;
		return;
	}
	orgjungsan_gubun = frm.jungsan_gubun.value;

	if (frm.jungsan_gubun.value == "����(�ؿ�)") {
		// �ؿܴ� ����ڹ�ȣ �ڵ������ȴ�(888-00-00000)

		//company_no.className = "text_ro";
		frm.company_no.readOnly = true;
		frm.company_no.value = "888-00-00000";
		$("#company_no").addClass("readonly");
		frm.checksocnoyn.value = "Y";
	} else {
		//company_no.className = "text";
		frm.company_no.readOnly = false;
		frm.company_no.value = "";
		$("#company_no").removeClass("readonly");
		frm.checksocnoyn.value = "N";
	}
}

function fnGetCheckCompanyNo() {
    var company_no = document.frmupche.company_no.value;
    if(company_no==""){
        alert("����� ��ȣ�� �Է����ּ���.");
    }
    else{
        var uid = document.frmupche.uid.value;
        $.ajax({
            type: "POST",
            url: "ajaxGetCommayNoCheck.asp",
            data: "company_no="+company_no + "&uid="+uid,
            cache: false,
            success: function(message) {
                if(message=="T") {
                    $("#pdiv").css("display","");
                    $("#company_no").attr("readonly",true);
                    $("#company_no").addClass("readonly");
                }else if(message=="C"){
                    alert("�̹� ���� �Է��� �Ϸ��Ͽ����ϴ�.");
                } else {
                    alert("��ϵ� ����ڹ�ȣ(�ֹε�Ϲ�ȣ)�� ��ġ���� �ʽ��ϴ�.\n��� MD���� �������ּ���.");
                }
            },
            error: function(err) {
                alert(err.responseText);
            }
        });
    }
}

// �̹������
function jsRegImg(sType, iMW, iMH, pvWidth){
    var winImg = window.open('/common/partner/popRegImg.asp?sType='+sType+'&iMH='+iMH+'&iMW='+iMW+'&pvWidth='+pvWidth,'popImg','width=500,height=350,scrollbars=yes,resizable=yes');
    winImg.focus();
}

function fnbizNoHyphen(num) {
     num = num.replace(/-/g, "");
     var num_str = num.toString();
     var result = '';
 
      for(var i=0; i<num_str.length; i++) {
            var tmp = num_str.length-(i+1);
            if(i==5){
				result = '-' + result;
			}
			else if(i==7){
				result = '-' + result;
			}
            result = num_str.charAt(tmp) + result;
       }
       return result;
}

function jsCheckUpload() {
    if($("#fileupload").val()!="") {
        $("#fileupmode").val("upload");
        $("#sType").val("company_no_img");
        $('#ajaxform').ajaxSubmit({
            //�������� validation check�� �ʿ��Ұ��
            beforeSubmit: function (data, frm, opt) {
                if(!(/\.(jpg|jpeg|png)$/i).test(frm[0].company_no_img2.value)) {
                    alert("JPG,PNG �̹������ϸ� ���ε� �Ͻ� �� �ֽ��ϴ�.");
                    $("#fileupload").val("");
                    return false;
                }
            },
            url: "<%=uploadImgUrl%>/linkweb/partnerAdmin/JoinUpload_Ajax2.asp",
            //submit������ ó��
            success: function(responseText, statusText){
                var resultObj = JSON.parse(responseText)
 
                if(resultObj.response=="fail") {
                    alert(resultObj.faildesc);
                } else if(resultObj.response=="ok") {
                    //document.frm1.bannerImg.value=resultObj.fileurl;
                    $("#company_no_img").val(resultObj.fileurl);
                    //alert(resultObj.fileurl);
                    $("#lyrBnrImg").hide().attr("src",$("#company_no_img").val()).fadeIn("fast");
                } else {
                    alert("ó���� ������ �߻��߽��ϴ�.\n" + resultObj.faildesc);
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

function jsCheckUpload2() {
    if($("#fileupload2").val()!="") {
        $("#fileupmode").val("upload");
        $("#sType").val("jungsan_acctno_img");
        
        $('#ajaxform').ajaxSubmit({
            //�������� validation check�� �ʿ��Ұ��
            beforeSubmit: function (data, frm, opt) {
                if(!(/\.(jpg|jpeg|png)$/i).test(frm[0].jungsan_acctno_img2.value)) {
                    alert("JPG,PNG �̹������ϸ� ���ε� �Ͻ� �� �ֽ��ϴ�.");
                    $("#fileupload2").val("");
                    return false;
                }
            },
            url: "<%=uploadImgUrl%>/linkweb/partnerAdmin/JoinUpload_Ajax3.asp",
            //submit������ ó��
            success: function(responseText, statusText){
                var resultObj = JSON.parse(responseText)
 
                if(resultObj.response=="fail") {
                    alert(resultObj.faildesc);
                } else if(resultObj.response=="ok") {
                    //document.frm1.bannerImg.value=resultObj.fileurl;
                    $("#jungsan_acctno_img").val(resultObj.fileurl);
                    //alert(resultObj.fileurl);
                    $("#lyrBnrImg2").hide().attr("src",$("#jungsan_acctno_img").val()).fadeIn("fast");
                } else {
                    alert("ó���� ������ �߻��߽��ϴ�.\n" + resultObj.faildesc);
                }
                $("#fileupload2").val("");
            },
            //ajax error
            error: function(err){
                alert("ERR: " + err.responseText);
                $("#fileupload2").val("");
            }
        });
    }
}

function fnPhoneNoHyphen(str, field){
    var str;
	str = checkDigit(str);
	len = str.length;
    if(len==8){
        if(str.substring(0,2)==02){
        error_numbr(str, field);
        }else{
        field.value = phone_format(1,str);
        }   
    }
    else if(len==9){
        if(str.substring(0,2)==02){
            field.value = phone_format(2,str);
        }
        else{
            error_numbr(str, field);
        }
    }
    else if(len==10){
        if(str.substring(0,2)==02){
            field.value = phone_format(2,str);
        }
        else{
            field.value = phone_format(3,str);
        }
    }
    else if(len==11){
        if(str.substring(0,2)==02){
        error_numbr(str, field);
        }else{
        field.value = phone_format(3,str);
        }
    }
    else{
        error_numbr(str, field);
    }
}

function checkDigit(num){
	var Digit = "1234567890";
	var string = num;
	var len = string.length
	var retVal = "";

	for (i = 0; i < len; i++){
		if (Digit.indexOf(string.substring(i, i+1)) >= 0){
			retVal = retVal + string.substring(i, i+1);
		}
	}
	return retVal;
}

function phone_format(type, num){
	if(type==1){
		return num.replace(/([0-9]{4})([0-9]{4})/,"$1-$2");
	}
	else if(type==2){
		return num.replace(/([0-9]{2})([0-9]+)([0-9]{4})/,"$1-$2-$3");
	}
	else{
		return num.replace(/(^01.{1}|[0-9]{3})([0-9]+)([0-9]{4})/,"$1-$2-$3");
	}
}

function error_numbr(str, field){
	alert("�������� ��ȣ�� �ƴմϴ�.");
	field.value = "";
	field.focus();
	return;
}
</script>
</head>

<div class="content scrl" style="top:0;">
    <div class="cont">
    <div class="popHead">
		<h1><img src="/images/partner/pop_admin_logo.gif" alt="10x10" /></h1>
	</div>
        <div class="pad20" style="max-width:640px;margin:0 auto;"> 
        <form name="frmupche" method="post" action="/common/partner/companyInfo_proc_new.asp">
        <input type="hidden" name="mode" value="addnewupchebrand" />
        <input type="hidden" name="uid" value="<%=pid%>" />
        <input type="hidden" name="pcuserdiv" value="<%=pcuserdiv%>" />
        <input type="hidden" name="checksocnoyn" value="N">
        <input type="hidden" name="company_no_img" id="company_no_img">
        <input type="hidden" name="jungsan_acctno_img" id="jungsan_acctno_img">
        <h3>�귣�� ����</h3>
        <table class="tbType1 writeTb tMar05">
            <colgroup>
                <col width="30%;" /><col width="70%" />
            </colgroup>
            <tbody>
            <tr>
                <th style="background-color:#FFE6E6;"><div>�����귣��ID</div></th>
                <td><%= pid %></td>
            </tr>
            <tr>
                <th style="background-color:#FFE6E6;"><div>�귣���(KR)</div></th>
                <td><% =opartner.FOneItem.Fsocname_kor %></td>
            </tr>
            <tr>
                <th style="background-color:#FFE6E6;"><div>�귣���(EN)</div></th>
                <td><% =opartner.FOneItem.Fsocname %></td>
            </tr>
            <tr>
                <th style="background-color:#FFE6E6;"><div>��������</div></th>
                <td><% =opartner.FOneItem.fpurchasetypename %></td>
            </tr>
            <tr>
                <th style="background-color:#FFE6E6;"><div>������</div></th>
                <td><% if opartner.FOneItem.Fjungsan_date<>"" then %>�¶��� : <% =opartner.FOneItem.Fjungsan_date %><% end if %><% if opartner.FOneItem.Fjungsan_date_off<>"" then %>�������� : <% =opartner.FOneItem.Fjungsan_date_off %><% end if %></td>
            </tr>
            <tr>
                <th style="background-color:#FFE6E6;"><div>�¶��� ī�װ�</div></th>
                <td><% SelectBoxBrandCategory "catecode2",opartner.FOneItem.Fcatecode %><input type="hidden" name="catecode" value="<%=opartner.FOneItem.Fcatecode%>"></td>
            </tr>
            <tr>
                <th style="background-color:#FFE6E6;"><div>���MD</div></th>
                <td><% drawSelectBoxCoWorker_OnOff "mduserid2", opartner.FOneItem.Fmduserid , "on" %><input type="hidden" name="mduserid" value="<%=opartner.FOneItem.Fmduserid%>"></td>
            </tr>
            <tr>
                <th style="background-color:#FFE6E6;"><div>�귣�� �⺻����</div></th>
                <td>
                <% =opartner.FOneItem.GetMWUName %>&nbsp;<% =opartner.FOneItem.Fdefaultmargine %>%&nbsp;&nbsp;
                <% if (opartner.FOneItem.FdefaultdeliveryType="9") then %>
                ��ü���� <%= FormatNumber(opartner.FOneItem.FDefaultFreeBeasongLimit,0) %>�� �̸� <%= FormatNumber(opartner.FOneItem.FDefaultDeliverPay,0) %>
                <% elseif (opartner.FOneItem.FdefaultdeliveryType="7") then %>
                ��ü����
                <% else %>
                �⺻��å (�ٹ�� : 3���� �̸� 2,000�� , ��ü��� : ����)
                <% end if %>
                </td>
            </tr>
            <tr>
                <th style="background-color:#FFE6E6;"><div>����ڹ�ȣ</div></th>
                <td colspan="3"><input type="text" class="formTxt" name="company_no" id="company_no" onkeyup="this.value=fnbizNoHyphen(this.value)" style="width:40%;" />&nbsp;<input type="button" class="btn3 btnIntb" value="Ȯ��" onclick="fnGetCheckCompanyNo()" /></td>
            </tr>
            </tbody>
        </table>
        </div>  
    </div>
    <div class="cont" id="pdiv" style="display:none">
        <div class="pad20" style="max-width:640px;margin:0 auto;"> 
        <h3>��Ʈ�� ����� ����</h3>
        
        <table class="tbType1 writeTb tMar05">
            <colgroup>
                <col width="30%;" /><col width="70%" />
            </colgroup>
            <tbody>
            <tr>
                <th><div>��й�ȣ 1��</div></th>
                <td><input type="text" class="formTxt" name="password" value="" size="20" maxlength="24"/></td>
            </tr>
            <tr>
                <th><div>��й�ȣ 2��</div></th>
                <td><input type="text" class="formTxt" name="passwordS" value="" size="20" maxlength="24"/></td>
            </tr>
            </tbody>
        </table><br>
        <span style="color:#ff0000; font-weight:bold;">* ����ڵ������ �Է½� ����ڵ������ ����� ����� �����ϰ� �Է����ּ���.</span>
        <h4 class="tMar20">��Ʈ�� ����ڵ������</h4>
        <table class="tbType1 writeTb">
            <colgroup>
                <col width="30%;" /><col width="70%" />
            </colgroup>
            <tbody>
            <tr>
                <th><div>ȸ���(��ȣ)</div></th>
                <td><input type="text" class="formTxt" name="company_name" value="" style="width:90%;" /></td>
            </tr>
            <tr>
                <th><div>��ǥ��</div></th>
                <td><input type="text" class="formTxt" name="ceoname" value="" style="width:90%;" /></td>
            </tr>
            <tr>
                <th><div>��������</div></th>
                <td>
                    <select name="jungsan_gubun" class="select" onchange="fnJungsanGubunChanged()">
                    <option value="�Ϲݰ���" >�Ϲݰ���</option>
                    <option value="���̰���" >���̰���</option>
                    <option value="��õ¡��" >��õ¡��</option>
                    <option value="�鼼" >�鼼</option>
                    <option value="����(�ؿ�)" >����(�ؿ�)</option>
                    </select>
                </td>
            </tr>
            <tr>
                <th><div>����������</div></th>
                <td> 
                    <p>
                        <input type="text" class="formTxt" name="company_zipcode" value="" maxlength="7" style="width:100px;" /> 
                        <input type="button" class="btn3 btnIntb" value="�˻�" onclick="FnFindZipNewPartner('frmupche','C')" />
                        <input type="button" class="btn3 btnIntb" value="�˻�(��)" onclick="TnFindZipPartnerDistribution('frmupche','C')" />
                        <% '<input type="button" class="btn3 btnIntb" value="�˻�(��)" onclick="javascript:popZip('s');" /> %>
                    </p>
                    <p class="tPad05"><input type="text" class="formTxt" name="company_address" value="" maxlength="64" style="width:96%;" /></p>
                    <p class="tPad05"><input type="text" class="formTxt" name="company_address2" value="" maxlength="64" style="width:96%;" /></p>
                </td>
            </tr>
            <tr>
                <th><div>����</div></th>
                <td><input type="text" class="formTxt" name="company_uptae" value="" maxlength="32" style="width:90%;" /></td>
            </tr>
            <tr>
                <th><div>����</div></th>
                <td><input type="text" class="formTxt" name="company_upjong" value="" maxlength="32" style="width:90%;" /></td>
            </tr>
            </tbody>
        </table>
        
        <h4 class="tMar20">��Ʈ�� ������������</h4>
        <table class="tbType1 writeTb">
            <colgroup>
                <col width="30%;" /><col width="70%" />
            </colgroup>
            <tbody>
            <tr>
                <th><div>�ŷ�����</div></th>
                <td>
                    <% DrawBankCombo "jungsan_bank", "" %>
                </td>
            </tr>
            <tr>
                <th><div>���¹�ȣ</div></th>
                <td><input type="text" class="formTxt" name="jungsan_acctno" value="" style="width:140px;" /></td>
            </tr>
            <tr>
                <th><div>�����ָ�</div></th>
                <td><input type="text" class="formTxt" name="jungsan_acctname" value="" style="width:140px;" /></td>
            </tr>
            </tbody>
        </table>
        
        <h4 class="tMar20">��Ʈ�� �⺻����</h4>
        <table class="tbType1 writeTb">
            <colgroup>
                <col width="30%" /><col width="70%;" />
            </colgroup>
            <tbody>
            <tr>
                <th><div>��ǥ��ȭ</div></th>
                <td><input type="text" class="formTxt" name="company_tel" value="" maxlength="16" style="width:90%;"  onfocusout="fnPhoneNoHyphen(this.value,this)"/></td>
            </tr>
            <tr>
                <th><div>�ѽ�</div></th>
                <td><input type="text" class="formTxt" name="company_fax" value="" maxlength="16" style="width:90%;"  onfocusout="fnPhoneNoHyphen(this.value,this)"/></td>
            </tr>
            <tr>
                <th><div>�繫���ּ�</div></td>
                <td> 
                    <p>
                        <input type="text" class="formTxt" name="return_zipcode" value="" maxlength="7" style="width:100px;" /> 
                        <input type="button" class="btn3 btnIntb" value="�˻�" onclick="FnFindZipNewPartner('frmupche','D')" /> 
                        <input type="button" class="btn3 btnIntb" value="�˻�(��)" onclick="TnFindZipPartnerDistribution('frmupche','D')" /> 
                        <% '<input type="button" class="btn3 btnIntb" value="�˻�(��)" onclick="javascript:popZip('m');" /> %>
                        <input type="checkbox" class="formCheck" name="samezip" id="samezip" onclick="SameReturnAddr(this.checked)" />
                        <label for="samezip">��</label>
                    </p>
                    <p class="tPad05"><input type="text" class="formTxt" name="return_address" value="" size="26" maxlength="64" style="width:96%;" /></p>
                    <p class="tPad05"><input type="text" class="formTxt" name="return_address2" value="" size="38" maxlength="64" style="width:96%;" /></p>
                </td>
            </tr>
            </tbody>
        </table>

        <h4 class="tMar20">��Ʈ�� ���������</h4>
        <table class="tbType1 writeTb">
            <colgroup>
                <col width="30%" /><col width="70%;" />
            </colgroup>
            <tbody>
            <tr>
                <th><div>����ڸ�</div></th>
                <td><input type="text" class="formTxt" name="manager_name" value="" maxlength="32" style="width:90%;" /></td>
            </tr>
            <tr>
                <th><div>�Ϲ���ȭ</div></th>
                <td><input type="text" class="formTxt" name="manager_phone" value="" maxlength="16" style="width:90%;" onfocusout="fnPhoneNoHyphen(this.value,this)"/></td>
            </tr>
            <tr>
                <th><div>E-Mail</div></th>
                <td><input type="text" class="formTxt" name="manager_email" value="" maxlength="64" style="width:90%;" /></td>
            </tr>
            <tr style="border-bottom:solid 2px #000;">
                <th><div>�ڵ���</div></th>
                <td><input type="text" class="formTxt" name="manager_hp" value="" maxlength="16" style="width:90%;" onfocusout="fnPhoneNoHyphen(this.value,this)"/></td>
            </tr>
            <tr>
                <th><div>�������ڸ�</div></th>
                <td><input type="text" class="formTxt" name="jungsan_name" value="" maxlength="32" style="width:90%;" /></td>
            </tr>
            <tr>
                <th><div>�Ϲ���ȭ</div></th>
                <td><input type="text" class="formTxt" name="jungsan_phone" value="" maxlength="16" style="width:90%;" onfocusout="fnPhoneNoHyphen(this.value,this)"/></td>
            </tr>
            <tr>
                <th><div>E-Mail</div></th>
                <td><input type="text" class="formTxt" name="jungsan_email" value="" maxlength="64" style="width:90%;" /></td>
            </tr>
            <tr>
                <th><div>�ڵ���</div></th>
                <td><input type="text" class="formTxt" name="jungsan_hp" value="" maxlength="16" style="width:90%;" onfocusout="fnPhoneNoHyphen(this.value,this)"/></td>
            </tr>
            <tbody>
        </table>
        <h4 class="tMar20">�귣�� ��۴�� �� ��ǰ�ּ�</h4>
        <table class="tbType1 writeTb">
            <colgroup>
                <col width="30%" /><col width="70%;" />
            </colgroup>
            <tbody>
            <tr>
                <th><div>��� ����ڸ�</div></th>
                <td><input type="text" class="formTxt" name="deliver_name" value="" maxlength="16" style="width:90%;" /></td>
            </tr>
            <tr>
                <th><div>�Ϲ���ȭ</div></th>
                <td><input type="text" class="formTxt" name="deliver_phone" value="" maxlength="16" style="width:90%;" onfocusout="fnPhoneNoHyphen(this.value,this)"/></td>
            </tr>
            <tr>
                <th><div>E-Mail</div></th>
                <td><input type="text" class="formTxt" name="deliver_email" value="" maxlength="64" style="width:90%;" /></td>
            </tr>
            <tr>
                <th><div>�ڵ���</div></th>
                <td><input type="text" class="formTxt" name="deliver_hp" value="" maxlength="16" style="width:90%;" onfocusout="fnPhoneNoHyphen(this.value,this)"/></td>
            </tr>
            <tr>
                <th><div>����(��ǰ)�ּ�</div></td>
                <td> 
                    <p>
                        <input type="text" class="formTxt" name="p_return_zipcode" value="" maxlength="7" style="width:100px;" /> 
                        <input type="button" class="btn3 btnIntb" value="�˻�" onclick="FnFindZipNewPartner('frmupche','I')" /> 
                        <input type="button" class="btn3 btnIntb" value="�˻�(��)" onclick="TnFindZipPartnerDistribution('frmupche','H')" /> 
                        <% '<input type="button" class="btn3 btnIntb" value="�˻�(��)" onclick="javascript:popZip('m');" /> %>
                        <input type="checkbox" class="formCheck" name="samezip" id="samezip" onclick="SameReturnAddr2(this.checked)" />
                        <label for="samezip">��</label>
                    </p>
                    <p class="tPad05"><input type="text" class="formTxt" name="p_return_address" value="" size="26" maxlength="64" style="width:96%;" /></p>
                    <p class="tPad05"><input type="text" class="formTxt" name="p_return_address2" value="" size="38" maxlength="64" style="width:96%;" /></p>
                </td>
            </tr>
            <tr>
                <th><div>����ù��</div></td>
                <td><% drawSelectBoxDeliverCompany "defaultsongjangdiv","" %></td>
            </tr>
            </tbody>
        </table>
        </form>
        <form name="frmUpload" id="ajaxform" method="post" enctype="multipart/form-data">
        <input type="hidden" name="mode" id="fileupmode" value="upload">
        <input type="hidden" name="sType" id="sType">
        <h4 class="tMar20">����ڵ����, ���� �纻 <span class="fs11">(�̹��� ����)</span></h4>
        <table class="tbType1 writeTb">
            <colgroup>
                <col width="30%;" /><col width="70%" />
            </colgroup>
            <tbody>
            <tr>
                <th><div>����ڵ���� �纻</div></td>
                <td> 
                    <p class="tPad05"><input type="file" class="formTxt" name="company_no_img2" id="fileupload" onchange="jsCheckUpload();" accept="image/*" style="width:80%;" /></p>
                    <div style="width:60px; height:60px; float:left; vertical-align:top; text-align:center;">
                        <img id="lyrBnrImg" src="" style="height:58px;"/>
                    </div>
                </td>
            </tr>
            <tr>
                <th><div>���� �纻</div></td>
                <td> 
                    <p class="tPad05"><input type="file" class="formTxt" name="jungsan_acctno_img2" id="fileupload2" onchange="jsCheckUpload2();" accept="image/*" style="width:80%;" />
                    <div style="width:60px; height:60px; float:left; vertical-align:top; text-align:center;">
                        <img id="lyrBnrImg2" src="" style="height:58px;"/>
                    </div>
                </td>
            </tr>
            </tbody>
        </table>
        </form>
        <div class="tMar15 ct">
            <input type="button" class="btn3 btnRd" value="��Ʈ������ ����" onclick="SaveUpcheInfo(frmupche);" />
        </div>
    </div>
</div>
</body>
</html>
<%
set opartner = Nothing
%>
<script>
$(function(){
    $("select[name='catecode2']").attr("disabled","disabled");
    $("select[name='mduserid2']").attr("disabled","disabled");
});
</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->
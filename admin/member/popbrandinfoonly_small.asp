<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 브랜드 정보
' History : 2009.04.17 서동석 생성
'			2022.10.12 한용민 수정(전시카테고리담당MD 추가)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<% '<!-- #include virtual="/lib/checkAllowIPWithLog.asp" --> %>
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/partners/partnerusercls.asp"-->
<!-- #include virtual="/lib/classes/partners/contractcls2013.asp"-->
<!-- include virtual="/lib/classes/cscenter/cs_aslistcls.asp"-->
<!-- #include virtual="/lib/classes/displaycate/displaycateCls.asp"-->
<%
dim itableWidth : itableWidth=600

dim ogroup,opartner,i, designer, groupid, prevmonthsocno, prevmonthgroupid, pcuserdiv
	designer = Replace(Trim(request("designer")), "'", "")

set opartner = new CPartnerUser
	opartner.FRectDesignerID = designer
	opartner.GetOnePartnerNUser

if opartner.FResultCount<=0 then
	Call Alert_return("존재하지 않는 브랜드 아이디입니다.")
	dbget.close()	:	response.End
end if

pcuserdiv = opartner.FOneItem.Fpcuserdiv
groupid = opartner.FOneItem.FGroupid

''prevmonthsocno = opartner.GetPrevMonthSocNO(designer)
prevmonthgroupid = opartner.GetPrevMonthGroupID(designer)

set ogroup = new CPartnerGroup
	ogroup.FRectGroupid = opartner.FOneItem.FGroupid
	ogroup.GetOneGroupInfo

dim ooffontract
set ooffontract = new COffContractInfo
	ooffontract.FRectDesignerID = designer
	ooffontract.GetOffMajorContractInfo

''기본 계약서 리스트
dim oDftContractList
set oDftContractList = new CPartnerContract
	oDftContractList.FRectGroupID = groupid
	oDftContractList.getRecentDefaultContract

dim onContractSheet, offContractSheet
set onContractSheet = new CPartnerContract
	onContractSheet.FRectGroupID = groupid
	onContractSheet.FRectMakerid = designer
	onContractSheet.getRecentAddContract(false)

set offContractSheet = new CPartnerContract
	offContractSheet.FRectGroupID = groupid
	offContractSheet.FRectMakerid = designer
	offContractSheet.getRecentAddContract(true)

dim returnsongjangStr
returnsongjangStr = returnsongjangStr + "10x10" & chr(9)
returnsongjangStr = returnsongjangStr + "(주)텐바이텐" & chr(9)
returnsongjangStr = returnsongjangStr + ogroup.FOneItem.FCompany_name  & chr(9)
returnsongjangStr = returnsongjangStr + ogroup.FOneItem.Fdeliver_phone  & chr(9)
returnsongjangStr = returnsongjangStr + ogroup.FOneItem.Fdeliver_hp  & chr(9)
returnsongjangStr = returnsongjangStr + replace(ogroup.FOneItem.Freturn_zipcode,"-","") & chr(9)
returnsongjangStr = returnsongjangStr + ogroup.FOneItem.Freturn_address  & chr(9)
returnsongjangStr = returnsongjangStr + ogroup.FOneItem.Freturn_address2  & chr(9)
returnsongjangStr = returnsongjangStr + "10x10 반품" & chr(9)
returnsongjangStr = returnsongjangStr + "반품상품" & chr(9)
returnsongjangStr = returnsongjangStr + opartner.FOneItem.FID

'dim OReturnAddr
'set OReturnAddr = new CCSReturnAddress
'OReturnAddr.FRectMakerid = designer
'OReturnAddr.GetBrandReturnAddress

'if (getPartnerCommCodeName("pcuserdiv",pcuserdiv)="") then
'    response.write "브랜드 구분 오류 - 관리자 문의요망"
'    response.end
'end if

'9999_02 >매입처(일반)
'9999_14 >매입처(아카데미)
'
'999_50 >제휴사(온라인)
'501_21 >직영점
'503_21 >기타매장
'900_21 >출고처(기타)

%>

<!--
returnsongjangStr = FormatDate(now(),"0000.00.00 00:00:00")
returnsongjangStr = Replace(returnsongjangStr,".","")
returnsongjangStr = Replace(returnsongjangStr,":","")
returnsongjangStr = Replace(returnsongjangStr," ","")
returnsongjangStr = returnsongjangStr & chr(9)
-->

<script type="text/javascript">

function copyComp(comp) {
	comp.focus()
	comp.select()
	therange=comp.createTextRange()
	therange.execCommand("Copy")
}

function CopyZip(flag,post1,post2,add,dong){
	var frm = eval(flag);

	frm.return_zipcode.value= post1 + "-" + post2;
	frm.return_address.value= add;
	frm.return_address2.value= dong;
}

function popZip(flag){
	var popwin = window.open("/lib/searchzip3.asp?target=" + flag,"searchzip3","width=460 height=240 scrollbars=yes resizable=yes");
	popwin.focus();
}

function SaveBrandInfo(frm){
/*
	if (frm.userdiv.value.length<1){
		alert('업체 구분을 선택하세요.');
		frm.userdiv.focus();
		return;
	}
*/

    var pcuserdiv = frm.pcuserdiv.value;

    <% if (ogroup.FOneItem.FGroupId="") then %>
    if ((pcuserdiv!="999_50")&&(pcuserdiv!="501_21")&&(pcuserdiv!="503_21")&&(pcuserdiv!="900_21")){
        alert('업체정보를 먼저 저장 하신후 브랜드정보를 저장 할 수 있습니다.(업체코드 없음)');
        return;
    }
    <% end if %>

    if (frm.pcuserdiv.value.length<1){
        alert('업체 구분이 정의되지 않았습니다. 관리자 문의요망.');
		return;
    }

//    if (frm.password.value.length<1){
//		alert('브랜드 패스워드를 입력하세요.');
//		frm.password.focus();
//		return;
//	}

	if (frm.socname_kor.value.length<1){
		alert('스트리트명(한글)을 입력하세요.');
		frm.socname_kor.focus();
		return;
	}

	if (frm.socname.value.length<1){
		alert('스트리트명(영문)을 입력하세요.');
		frm.socname.focus();
		return;
	}

	if (frm.prtidx.value.length<1){
		alert('랙 번호를 입력하세요. - [기본값 9999]');
		frm.prtidx.focus();
		return;
	} else {
		if (frm.prtidx.value.length<4) {
			var cnt = parseInt(4-frm.prtidx.value.length);
			var tmpPrtidx;
			for(var i=0;i<cnt;i++){
				frm.prtidx.value='0' + frm.prtidx.value;
			}
		}
	}

    //일반 매입처.
    if (pcuserdiv=="9999_02"){
    	if ((!frm.isusing[0].checked)&&(!frm.isusing[1].checked)){
    		alert('사용여부를 선택하세요.');
    		frm.isusing[0].focus();
    		return;
    	}

        if ((!frm.isoffusing[0].checked)&&(!frm.isoffusing[1].checked)){
    		alert('사용여부를 선택하세요.');
    		frm.isoffusing[0].focus();
    		return;
    	}

		/*
		// 제휴몰 판매설정 팝업창에서 수정가능(skyer9)
    	if ((!frm.isextusing[0].checked)&&(!frm.isextusing[1].checked)){
    		alert('제휴몰 사용여부를 선택하세요.');
    		frm.isextusing[0].focus();
    		return;
    	}

        //제휴사 브랜드 설정 confirm 텐바이텐Y,제휴N인경우 (생각없이 N로 설정하므로 컨펌)
        if ((frm.isusing[0].checked)&&(frm.isextusing[1].checked)){
            if (!confirm('제휴몰 브랜드 사용여부 N인경우 InterPark,Lotte 등 제휴몰에 판매하지 않습니다. 계속하시겠습니까?')) {
                frm.isextusing[0].focus();
                return;
            }
        }
		*/

    	if ((!frm.streetusing[0].checked)&&(!frm.streetusing[1].checked)){
    		alert('스트리트 사용여부를 선택하세요.');
    		frm.streetusing[0].focus();
    		return;
    	}

		/*
    	if ((!frm.extstreetusing[0].checked)&&(!frm.extstreetusing[1].checked)){
    		alert('제휴몰 스트리트 사용여부를 선택하세요.');
    		frm.extstreetusing[0].focus();
    		return;
    	}
    	*/

    	if ((!frm.specialbrand[0].checked)&&(!frm.specialbrand[1].checked)){
    		alert('커뮤니티 사용여부를 선택하세요.');
    		frm.specialbrand[0].focus();
    		return;
    	}

    	if (frm.catecode.value.length<1 && frm.offcatecode.value.length<1){
    		alert('온라인 또는 오프라인 카테고리 구분을 선택하세요. \n- 둘 중 하나는 필수 사항입니다.');
    		//frm.catecode.focus();
    		return;
    	}

		if (frm.standardmdcatecode.value.length<1){
			alert('전시카테고리 담당MD를 선택하세요.');
			frm.standardmdcatecode.focus();
			return;
		}

        if (frm.mduserid.value.length<1 && frm.offmduserid.value.length<1){
    		alert('온라인 또는 오프라인 담당MD를 선택하세요. \n- 둘 중 하나는 필수 사항입니다.');
    		//frm.mduserid.focus();
    		return;
    	}

    	if (frm.maeipdiv.value.length<1){
    		alert('매입 구분을 선택하세요.');
    		frm.maeipdiv.focus();
    		return;
    	}

    	if (!IsDouble(frm.defaultmargine.value)){
    		alert('기본마진을 입력하세요. - 실수만 가능합니다.');
    		frm.defaultmargine.focus();
    		return;
    	}

        var addMargin = eval('document.'+frm.name+'.'+frm.maeipdiv.value+'_margin');
        if (addMargin){
            if (addMargin.value!=frm.defaultmargine.value){
                //2014/01/01 부터 적용
                //alert('기본마진과 추가마진이 일치 하지 않습니다.\n\n향후 추가 마진을 사용할 예정이므로 일치시켜 주시기 바랍니다.');
                //addMargin.focus();
                //return;
                if (!confirm('기본마진과 추가마진이 일치 하지 않습니다.\n\n계속 하시겠습니까?')){
                    addMargin.focus();
                    return;
                }
            }
        }

    	//조건배송 수정은 팝업창에서;
    }

	//매입처(아카데미)
    if (pcuserdiv=="9999_14"){
        var selltype=frm.selltype.value;

        if (frm.mduserid.value.length<1){
    		alert('담당MD 구분을 선택하세요.');
    		frm.mduserid.focus();
    		return;
    	}

        var lec_yn = getFieldValue(frm.lec_yn);
        var diy_yn = getFieldValue(frm.diy_yn);

        if ((lec_yn=="N")&&(diy_yn=="N")){
            alert('강좌/DIY 둘중 하나는 사용으로 설정하셔야 합니다.');
            frm.lec_yn[0].focus();
            return;
        }

        if ((lec_yn=="Y")&&(frm.lec_margin.value.length<1)){
            alert('강좌 기본 마진을 입력하세요.');
            frm.lec_margin.focus();
            return;
        }

        if ((lec_yn=="Y")&&(frm.mat_margin.value.length<1)){
            alert('재료비 기본 마진을 입력하세요.');
            frm.mat_margin.focus();
            return;
        }

        if ((diy_yn=="Y")&&(frm.diy_margin.value.length<1)){
            alert('DIY상품 기본 마진을 입력하세요.');
            frm.diy_margin.focus();
            return;
        }

        if ((frm.diy_yn[0].checked)&&(frm.diy_dlv_gubun.value.length<1)){
            alert('DIY 배송구분을 선택하세요.');
            frm.diy_dlv_gubun.focus();
            return;
        }

        if (frm.diy_dlv_gubun.value=="9"){
            if (!IsDigit(frm.DefaultFreebeasongLimit.value)){
                alert('배송비 기준 숫자만 가능합니다.');
                frm.DefaultFreebeasongLimit.focus();
                return;
            }

            if (!IsDigit(frm.DefaultDeliverPay.value)){
                alert('배송비  숫자만 가능합니다.');
                frm.DefaultDeliverPay.focus();
                return;
            }

            if (frm.DefaultFreebeasongLimit.value*1<=0){
                alert('금액을 0원 이상 입력하세요.');
                frm.DefaultFreebeasongLimit.focus();
                return;
            }

            if (frm.DefaultDeliverPay.value*1<=0){
                alert('금액을 0원 이상 입력하세요.');
                frm.DefaultDeliverPay.focus();
                return;
            }

        }

        if ((lec_yn=="Y")&&(diy_yn=="N")){
            frm.selltype.value="10";
            frm.maeipdiv.value="M";
            frm.defaultmargine.value=frm.lec_margin.value;
        }

        if ((lec_yn=="N")&&(diy_yn=="Y")){
            frm.selltype.value="20";
            frm.maeipdiv.value="U";
            frm.defaultmargine.value=frm.diy_margin.value;

        }

        if ((lec_yn=="Y")&&(diy_yn=="Y")){
            frm.selltype.value="30";

        }

        if (diy_yn=="Y"){
            frm.maeipdiv.value="U";
            frm.defaultmargine.value=frm.diy_margin.value;
            frm.defaultdeliverytype.value = frm.diy_dlv_gubun.value;
        }
	}

	// 온라인제휴사, etc출고처
	if ((pcuserdiv=="999_50")||(pcuserdiv=="900_21")){
	    if (frm.purchasetype.value.length<1){
	        alert('정산 방식을 선택 하세요. 필수 값입니다.');
	        frm.purchasetype.focus()
	        return;
	    }
	}

	//온라인제휴사
	if (pcuserdiv=="999_50"){
	    if (frm.commission.value.length<1){
	        alert('수수료를 입력 하세요.');
	        frm.commission.focus()
	        return;
	    }
	}

	//스트리트 표시여부 제휴몰을 텐바이텐과 통일.
	if (pcuserdiv=="9999_02"){
    	if(frm.streetusing[0].checked){
    		frm.extstreetusing.value = "Y";
    	}else if(frm.streetusing[1].checked){
    		frm.extstreetusing.value = "N";
    	}
    }

    if (frm.tplcompanyid){
        if (frm.tplcompanyid.value.length>0){
            if (!confirm('3pl 연결계정이 선택 되었습니다. 계속하시겠습니까?')){
                return;
            }
        }
    }

	var ret = confirm('브랜드 정보를 저장 하시겠습니까?');

	if (ret){
		frm.submit();
	}
}

function SaveUpcheInfo(frm){
	if (frm.groupid.value.length<1){
		var ret = confirm('업체 정보를 저장 하시겠습니까?');
	}else{
		var ret = confirm('같은 그룹코드에 있는 기존 업체 정보도 수정됩니다. 저장 하시겠습니까?');
	}

	if (ret){
		frm.submit();
	}
}

function ModiInfo(frm){
	var ret = confirm('저장 하시겠습니까?');

	if (ret){
		//frm.submit();
	}

}

function CopyFromBrandInfo(){
	frmupche.company_name.value = frmbuf.company_name.value;
	frmupche.ceoname.value = frmbuf.ceoname.value;
	frmupche.company_no.value = frmbuf.company_no.value;
	frmupche.jungsan_gubun.value = frmbuf.jungsan_gubun.value;
	frmupche.company_zipcode.value = frmbuf.company_zipcode.value;
	frmupche.company_address.value = frmbuf.company_address.value;
	frmupche.company_address2.value = frmbuf.company_address2.value;
	frmupche.company_uptae.value = frmbuf.company_uptae.value;
	frmupche.company_upjong.value = frmbuf.company_upjong.value;
	frmupche.company_tel.value = frmbuf.company_tel.value;
	frmupche.company_fax.value = frmbuf.company_fax.value;
	frmupche.jungsan_bank.value = frmbuf.jungsan_bank.value;
	frmupche.jungsan_acctno.value = frmbuf.jungsan_acctno.value;
	frmupche.jungsan_acctname.value = frmbuf.jungsan_acctname.value;
	frmupche.manager_name.value = frmbuf.manager_name.value;
	frmupche.manager_phone.value = frmbuf.manager_phone.value;
	frmupche.manager_email.value = frmbuf.manager_email.value;
	frmupche.manager_hp.value = frmbuf.manager_hp.value;

	frmupche.deliver_name.value = frmbuf.deliver_name.value;
	frmupche.deliver_phone.value = frmbuf.deliver_phone.value;
	frmupche.deliver_email.value = frmbuf.deliver_email.value;
	frmupche.deliver_hp.value = frmbuf.deliver_hp.value;

	frmupche.jungsan_name.value = frmbuf.jungsan_name.value;
	frmupche.jungsan_phone.value = frmbuf.jungsan_phone.value;
	frmupche.jungsan_email.value = frmbuf.jungsan_email.value;
	frmupche.jungsan_hp.value = frmbuf.jungsan_hp.value;
}

function PopUpcheList(frmname){
	var popwin = window.open("/admin/lib/popupchelist.asp?frmname=" + frmname,"popupchelist","width=700 height=540 scrollbars=yes resizable=yes");
	popwin.focus();
}

function SaveBrandEtcInfo(frm){
//	if (!FileCheck(frm.logoimg,150000,160,110)){
//		frm.file1.focus();
//		return;
//	}
//
//	if (!FileCheck(frm.titleimg,1500000,720,220)){
//		frm.file2.focus();
//		return;
//	}

	if (confirm('저장 하시겠습니까?')){
		frm.submit();
	}

}

function ChangeTitle(comp,imgcomp){
	imgcomp.src = comp.value;
}

function ChangeLogo(comp,imgcomp){
	imgcomp.src = comp.value;
}

function ChangeIcon(comp,imgcomp){
	imgcomp.src = comp.value;
}

function FileCheck(comp,maxfilesize,maxwidth,maxheight){
	if(comp.fileSize > maxfilesize){
		alert("파일사이즈는 "+ maxfilesize + "byte를 넘기실 수 없습니다...");
		return false;
	}

	if ((comp.src!="")&&(comp.width <1)){
		alert('이미지만 가능합니다.');
		return false;
	}

	//if(comp.width > maxwidth){
	//	alert("가로폭은 " + maxwidth + " 픽셀을 넘기실 수 없습니다...");
	//	return false;
	//}
	//if(comp.height > maxheight){
	//	alert("세로폭은 " + maxheight + " 픽셀을 넘기실 수 없습니다...");
	//	return false;
	//}

	return true;
}

function MakeReturnSongjang(makerid, chargeArrival) {
    var paramgubunname = "반품";
	if (chargeArrival == 'Y') {
		paramgubunname = paramgubunname + '(착불)';
	}
    if (confirm(paramgubunname + ' 송장을 생성 하시겠습니까?')){
    	var popwin = window.open("/common/action/popbrandsongjangMake.asp?makerid=" + makerid + "&paramgubunname=" + paramgubunname + "&chargeArrival=" + chargeArrival,"popbrandsongjang","width=100 height=100 scrollbars=yes resizable=yes");
    	popwin.focus();
	}
}

function RegContract(makerid,groupid){
    var popwin = window.open('/admin/member/contract/ctrReg.asp?makerid=' + makerid+'&groupid='+groupid,'ctrReg','width=1124,height=768,scrollbars=yes,resizable=yes');
    popwin.focus();
}

function ViewContract(makerid,ContractID){
    var popwin = window.open('contractReg.asp?makerid=' + makerid + '&ContractID='+ContractID,'contractView','width=860,height=600,scrollbars=yes,resizable=yes');
    popwin.focus();
}

function chkCompdiygbn(comp){
    var frm = comp.form;
    if (comp.value=="9"){
        frm.DefaultFreebeasongLimit.style.background = '#FFFFFF';
        frm.DefaultDeliverPay.style.background  = '#FFFFFF';

        frm.DefaultFreebeasongLimit.readOnly = false;
        frm.DefaultDeliverPay.readOnly = false;

        frm.DefaultFreebeasongLimit.value=frm.pDFL.value;
        frm.DefaultDeliverPay.value=frm.pDDP.value;


    }else{
        frm.DefaultFreebeasongLimit.style.background = '#BBBBBB';
        frm.DefaultDeliverPay.style.background  = '#BBBBBB';

        frm.DefaultFreebeasongLimit.readOnly = true;
        frm.DefaultDeliverPay.readOnly = true;

        frm.DefaultFreebeasongLimit.value=0;
        frm.DefaultDeliverPay.value=0;
    }
}

function clickLec(comp){

}

function clickDiy(comp){
    if (comp.value=="Y"){
        iDiyDlv.style.display="";
    }else{
        iDiyDlv.style.display="none";
    }
}

function popShopInfo(ishopid){
	var popwin = window.open("/admin/lib/popoffshopinfo.asp?shopid=" + ishopid + "&menupos=277","popoffshopinfo",'width=1024,height=768,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function popEtcSiteUsing(){
    var popwin = window.open("/admin/etc/outmall/popJaehyu_Not_In_Makerid.asp?isBrandPage=Y&makerid=<%=designer%>","popEtcSiteUsing",'width=1200,height=600,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function jsDispCate(){
    var popwin = window.open("/admin/member/popbrandinfoonly_dispcate.asp?makerid=<%=designer%>","popDispCate",'width=420,height=200,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function chgAddMargin(comp){
    var frm = comp.form;
    var tgcomp = eval('document.'+frm.name+'.'+frm.maeipdiv.value+'_margin');
    if (tgcomp){
        tgcomp.value = comp.value;
    }
}

//비밀번호 변경
function jsOpenPW(brandid,stype){
	var winPw = window.open("/admin/member/popbrandChangePW.asp?bid="+brandid+"&sT="+stype,"popPW","width=400, height=400,scrollbars=yes,resizable=yes");
	winPw.focus();
}

function jsModiPrevMonthGroupID(makerid) {
	var frm = document.frmbrand;

	if (confirm("브랜드(" + makerid + ")의 전월 그룹코드를 현재 그룹코드로 변경합니다.\n\n진행하시겠습니까?") == true) {
		frm.mode.value = "modiprevmonthgroupid";
		frm.submit();
	}
}

</script>
<form name="frmbuf" style="margin:0px;">
<input type=hidden name=company_name value="<%= opartner.FOneItem.FCompany_name %>">
<input type=hidden name=ceoname value="<%= opartner.FOneItem.Fceoname %>">
<input type=hidden name=company_no value="<%= socialnoBlank(opartner.FOneItem.Fcompany_no) %>">
<input type=hidden name=jungsan_gubun value="<%= opartner.FOneItem.Fjungsan_gubun %>">
<input type=hidden name=company_zipcode value="<%= opartner.FOneItem.Fzipcode %>">
<input type=hidden name=company_address value="<%= opartner.FOneItem.Faddress %>">
<input type=hidden name=company_address2 value="<%= opartner.FOneItem.Fmanager_address %>">
<input type=hidden name=company_uptae value="<%= opartner.FOneItem.Fcompany_uptae %>">
<input type=hidden name=company_upjong value="<%= opartner.FOneItem.Fcompany_upjong %>">
<input type=hidden name=company_tel value="<%= opartner.FOneItem.Ftel %>">
<input type=hidden name=company_fax value="<%= opartner.FOneItem.Ffax %>">

<input type=hidden name=jungsan_bank value="<%= opartner.FOneItem.Fjungsan_bank %>">
<input type=hidden name=jungsan_acctno value="<%= opartner.FOneItem.Fjungsan_acctno %>">
<input type=hidden name=jungsan_acctname value="<%= opartner.FOneItem.Fjungsan_acctname %>">
<input type=hidden name=manager_name value="<%= opartner.FOneItem.Fmanager_name %>">
<input type=hidden name=manager_phone value="<%= opartner.FOneItem.Fmanager_phone %>">
<input type=hidden name=manager_email value="<%= opartner.FOneItem.Femail %>">
<input type=hidden name=manager_hp value="<%= opartner.FOneItem.Fmanager_hp %>">

<input type=hidden name=deliver_name value="<%= opartner.FOneItem.Fdeliver_name %>">
<input type=hidden name=deliver_phone value="<%= opartner.FOneItem.Fdeliver_phone %>">
<input type=hidden name=deliver_email value="<%= opartner.FOneItem.Fdeliver_email %>">
<input type=hidden name=deliver_hp value="<%= opartner.FOneItem.Fdeliver_hp %>">

<input type=hidden name=jungsan_name value="<%= opartner.FOneItem.Fjungsan_name %>">
<input type=hidden name=jungsan_phone value="<%= opartner.FOneItem.Fjungsan_phone %>">
<input type=hidden name=jungsan_email value="<%= opartner.FOneItem.Fjungsan_email %>">
<input type=hidden name=jungsan_hp value="<%= opartner.FOneItem.Fjungsan_hp %>">
</form>
<table width="<%= itableWidth %>" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="" style="margin:0px;">
<input type="hidden" name="page" value="1">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="4">
		브랜드 ID : <input type="text" class="text" name="designer" value="<%= designer %>" Maxlength="32" size="16">
		<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
	</td>
</tr>
</form>

<form name="frmbrand" method="post" action="doupcheedit.asp" style="margin:0px;">
<input type="hidden" name="uid" value="<%= opartner.FOneItem.FID %>">
<input type="hidden" name="mode" value="brandedit">
<input type="hidden" name="pcuserdiv" value="<%=pcuserdiv%>">
<input type="hidden" name="groupid" value="<%= ogroup.FOneItem.FGroupId %>">
<tr height="25">
	<td bgcolor="<%= adminColor("pink") %>" colspan="4"><b>* 브랜드 기본정보(공통)</b></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td height="25" colspan="4">**브랜드 기본정보**</td>
</tr>
<tr >
	<td bgcolor="<%= adminColor("pink") %>">브랜드구분</td>
	<td bgcolor="#FFFFFF" >
	<%= getPartnerCommCodeName("pcuserdiv",pcuserdiv) %>

	</td>
	<td bgcolor="<%= adminColor("pink") %>">등록일</td>
	<td bgcolor="#FFFFFF" ><%= opartner.FOneItem.Fregdate %></td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("pink") %>">브랜드ID</td>
	<td bgcolor="#FFFFFF" colspan="3"><%= opartner.FOneItem.FID %>
		<% if C_ADMIN_AUTH or C_CSPowerUser then %><span style="padding-left:10px;"><input type="button" class="button" value="비밀번호 변경" onClick="jsOpenPW('<%=opartner.FOneItem.FID%>','P');"></span><%END IF%>
	</td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("pink") %>">브랜드명(KR)</td>
	<td bgcolor="#FFFFFF" >
	<input type="text" class="text" name="socname_kor" value="<%= opartner.FOneItem.Fsocname_kor %>">
	</td>
	<td bgcolor="<%= adminColor("pink") %>">브랜드명(EN)</td>
	<td bgcolor="#FFFFFF" >
	<input type="text" class="text" name="socname" value="<%= opartner.FOneItem.Fsocname %>">
	</td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("pink") %>" ></td>
	<td bgcolor="#FFFFFF" >

	</td>
	<td bgcolor="<%= adminColor("pink") %>">어드민오픈여부</td>
	<td bgcolor="#FFFFFF" >
		<b><%= fnColor(opartner.FOneItem.Fpartnerusing,"yn") %></b>
		&nbsp;&nbsp;
		<% if C_ADMIN_AUTH or C_OFF_AUTH or C_MD_AUTH then %>
		<input type="button" class="button" value="수정" onClick="PopBrandAdminUsingChange('<%= opartner.FOneItem.FID %>');">
		<font color="blue"><b>*</b></font>
		<% end if %>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td height="25" colspan="4">**브랜드 물류정보**</td>
</tr>
<tr >
    <td bgcolor="<%= adminColor("pink") %>">택배사</td>
	<td bgcolor="#FFFFFF" ><% drawSelectBoxDeliverCompany "defaultsongjangdiv",opartner.FOneItem.Fdefaultsongjangdiv %>
	<%= opartner.FOneItem.Ftakbae_tel %>
	</td>

	<td width="90" bgcolor="<%= adminColor("pink") %>" >랙번호(물류)</td>
	<td bgcolor="#FFFFFF" >
	<input type="text" name="prtidx" value="<%= opartner.FOneItem.getRackCode %>" size="4" maxlength="4">
	(기본값 : 9999)</td>
	</td>
</tr>
<tr height=25"">
    <td bgcolor="<%= adminColor("pink") %>"></td>
	<td bgcolor="#FFFFFF" ></td>

	<td width="90" bgcolor="<%= adminColor("pink") %>" >랙박스수량</td>
	<td bgcolor="#FFFFFF" >
		<%= opartner.FOneItem.Frackboxno %>
	</td>
	</td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("pink") %>" >반품(물류)주소</td>
	<td bgcolor="#FFFFFF" colspan=3 >
		<input type="text" class="text" name="return_zipcode" value="<%= opartner.FOneItem.Freturn_zipcode %>" size="7" maxlength="7">
	    <input type="button" class="button" value="검색" onClick="FnFindZipNew('frmbrand','D')" style="width: 50px;">
		<input type="button" class="button" value="검색(구)" onClick="TnFindZipNew('frmbrand','D')" style="width: 50px;">
	    <% '<input type="button" class="button" value="검색(구)" onClick="popZip('frmbrand');" style="width: 60px;"> %>
		&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
		<input type="button" class="button" value="일반 송장등록" onClick="MakeReturnSongjang('<%= opartner.FOneItem.FID %>', 'N');" style="width: 90px;">
		&nbsp;&nbsp;
		<input type="button" class="button" value="착불 송장등록" onClick="MakeReturnSongjang('<%= opartner.FOneItem.FID %>', 'Y');" style="width: 90px;">
		<br>
		<input type="text" class="text" name="return_address" value="<%= opartner.FOneItem.Freturn_address %>" size="25" maxlength="64">
		<input type="text" class="text" name="return_address2" value="<%= opartner.FOneItem.Freturn_address2 %>" size="40" maxlength="128">

	</td>
</tr>
<tr >
    <td bgcolor="<%= adminColor("pink") %>" height="25">배송담당자</td>
	<td bgcolor="#FFFFFF" >
		<input type="text" class="text" name="deliver_name" value="<%= opartner.FOneItem.Fdeliver_name %>" size="24" maxlength="32">
	</td>
	<td bgcolor="<%= adminColor("pink") %>">배송전화</td>
	<td bgcolor="#FFFFFF" >
		<input type="text" class="text" name="deliver_phone" value="<%= opartner.FOneItem.Fdeliver_phone %>" size="16" maxlength="16">
	</td>
</tr>
<tr >

    <td bgcolor="<%= adminColor("pink") %>">배송이메일</td>
	<td bgcolor="#FFFFFF" >
		<input type="text" class="text" name="deliver_email" value="<%= opartner.FOneItem.Fdeliver_email %>" size="24" maxlength="128">
	</td>
    <td bgcolor="<%= adminColor("pink") %>" height="25">배송핸드폰</td>
	<td bgcolor="#FFFFFF" >
		<input type="text" class="text" name="deliver_hp" value="<%= opartner.FOneItem.Fdeliver_hp %>" size="16" maxlength="16">
	</td>
</tr>
</table>

<br>

<% ''' 9999_15 추가 2016/05/16 %>
<% if (pcuserdiv="9999_02") or (pcuserdiv="9999_15") then %>
	<table width="<%= itableWidth %>" align="center" cellpadding="1" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr bgcolor="#FFFFFF">
		<% if (pcuserdiv="9999_15") then %>
	    <td height="25" colspan="6">**매입처(핑거스상품) 추가정보**</td>
	    <% else %>
		<td height="25" colspan="6">**매입처(일반) 추가정보**</td>
	    <% end if %>

	</tr>
	<tr>

		<td rowspan="3" bgcolor="<%= adminColor("pink") %>">브랜드<br>사용여부<br>(카테고리노출)</td>
		<td width="70" bgcolor="#FFFFFF">텐바이텐</td>
		<td bgcolor="#FFFFFF"><input type=radio name="isusing" value="Y" <% if opartner.FOneItem.Fisusing="Y" then response.write "checked" %> >Y <input type=radio name="isusing" value="N" <% if opartner.FOneItem.Fisusing="N" then response.write "checked" %> >N</td>

		<td width="100" bgcolor="<%= adminColor("pink") %>">구매유형</td>
		<td bgcolor="#FFFFFF" colspan=2>
			<% drawPartnerCommCodeBox false,"purchasetype","purchasetype",CHKIIF(opartner.FOneItem.FpurchaseType="","1",opartner.FOneItem.FpurchaseType),"" %>
		</td>
	</tr>
	<tr>
	    <td bgcolor="#FFFFFF">텐바이텐 OFF</td>
		<td bgcolor="#FFFFFF">
			<input type=radio name="isoffusing" value="Y" <% if opartner.FOneItem.Fisoffusing="Y" then response.write "checked" %> >Y <input type=radio name="isoffusing" value="N" <% if opartner.FOneItem.Fisoffusing="N" then response.write "checked" %> >N

		</td>
		<td rowspan="2" bgcolor="<%= adminColor("pink") %>">스트리트<br>표시여부<br>(브랜드운영관련)</td>
		<td bgcolor="#FFFFFF">텐바이텐</td>
		<td bgcolor="#FFFFFF"><input type=radio name="streetusing" value="Y" <% if opartner.FOneItem.Fstreetusing="Y" then response.write "checked" %> >Y <input type=radio name="streetusing" value="N" <% if opartner.FOneItem.Fstreetusing="N" then response.write "checked" %> >N</td>
	</tr>
	<tr >
		 <td bgcolor="#FFFFFF">제휴몰</td>
		<td bgcolor="#FFFFFF">
			<%= opartner.FOneItem.Fisextusing %> (팝업에서 수정가능)
			<!--
			<input type=radio name="isextusing" value="Y" <% if opartner.FOneItem.Fisextusing="Y" then response.write "checked" %> >Y <input type=radio name="isextusing" value="N" <% if opartner.FOneItem.Fisextusing="N" then response.write "checked" %> >N
			-->

			<!-- 스트리트 표시여부 제휴몰 히든처리. 스트리트 표시여부 텐바이텐 과 값 동일. 아래부분 기존소스. //-->
			<input type="hidden" name="extstreetusing" value="">
		</td>
		<td bgcolor="#FFFFFF">커뮤니티(상품Q/A)</td>
		<td bgcolor="#FFFFFF"><input type=radio name="specialbrand" value="Y" <% if opartner.FOneItem.Fspecialbrand="Y" then response.write "checked" %>>Y <input type=radio name="specialbrand" value="N" <% if opartner.FOneItem.Fspecialbrand="N" then response.write "checked" %>>N</td>
		<!--
		<td bgcolor="#FFFFFF">제휴몰</td>
		<td bgcolor="#FFFFFF"><input type=radio name="extstreetusing" value="Y" <% if opartner.FOneItem.Fextstreetusing="Y" then response.write "checked" %> >Y <input type=radio name="extstreetusing" value="N" <% if opartner.FOneItem.Fextstreetusing="N" then response.write "checked" %> >N	</td>
		//-->
	</tr>
	<tr >

		<td width="100" bgcolor="<%= adminColor("pink") %>" ><!-- 판매채널 --></td>
		<td bgcolor="#FFFFFF" colspan=2>
		<input class="button" type="button" value="제휴몰별 판매제외설정" onClick="popEtcSiteUsing();">
		</td>
		<td bgcolor="#FFFFFF" colspan="3"></td>

	</tr>
	<tr>
		<td bgcolor="<%= adminColor("pink") %>">Only 노출</td>
		<td bgcolor="#FFFFFF" colspan="2"><input type=radio name="onlyflg" value="Y" <% if opartner.FOneItem.Fonlyflg="Y" then response.write "checked" %> >Y <input type=radio name="onlyflg" value="N" <% if opartner.FOneItem.Fonlyflg="N" then response.write "checked" %> >N</td>
		<td bgcolor="<%= adminColor("pink") %>">Artist 노출</td>
		<td bgcolor="#FFFFFF" colspan="2"><input type=radio name="artistflg" value="Y" <% if opartner.FOneItem.Fartistflg="Y" then response.write "checked" %> >Y <input type=radio name="artistflg" value="N" <% if opartner.FOneItem.Fartistflg="N" then response.write "checked" %> >N</td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("pink") %>">K-Design 노출</td>
		<td bgcolor="#FFFFFF" colspan="2"><input type=radio name="kdesignflg" value="Y" <% if opartner.FOneItem.Fkdesignflg="Y" then response.write "checked" %> >Y <input type=radio name="kdesignflg" value="N" <% if opartner.FOneItem.Fkdesignflg="N" then response.write "checked" %> >N</td>
		<td bgcolor="#FFFFFF" colspan="3"></td>
	</tr>


	<tr >
		<td bgcolor="<%= adminColor("pink") %>">온라인 카테고리</td>
		<td bgcolor="#FFFFFF" colspan=2><% SelectBoxBrandCategory "catecode", opartner.FOneItem.Fcatecode %></td>
		<td bgcolor="<%= adminColor("pink") %>" >온라인 담당MD</td>
		<td bgcolor="#FFFFFF" colspan=2><% drawSelectBoxCoWorker_OnOff "mduserid", opartner.FOneItem.Fmduserid, "on" %></td>
	</tr>
	<tr >
		<td bgcolor="<%= adminColor("pink") %>">온라인 전시카테고리</td>
		<td bgcolor="#FFFFFF" colspan=2><input type="button" class="button" value="설정하기" onClick="jsDispCate()">
		<%= Chkiif(opartner.FOneItem.FstandardCateCode <> "","대표 전시카테고리 설정됨","") %>
		</td>
		<td bgcolor="<%= adminColor("pink") %>" >전시카테고리<br>담당MD</td>
		<td bgcolor="#FFFFFF" colspan=2><%= fnStandardDispCateSelectBox(1,"", "standardmdcatecode", opartner.FOneItem.Fstandardmdcatecode, "")%></td>
	</tr>
	<tr >
		<td bgcolor="<%= adminColor("pink") %>">오프라인 카테고리</td>
		<td bgcolor="#FFFFFF" colspan=2><% SelectBoxBrandCategory "offcatecode", opartner.FOneItem.Foffcatecode %></td>
		<td bgcolor="<%= adminColor("pink") %>" >오프라인 담당MD</td>
		<td bgcolor="#FFFFFF" colspan=2><% drawSelectBoxCoWorker_OnOff "offmduserid", opartner.FOneItem.Foffmduserid, "off" %></td>
	</tr>
	</table>

	<br>

	<table width="<%= itableWidth %>" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr bgcolor="#FFFFFF">
		<td colspan="6">**계약관련사항** </td>
	</td>
	<tr >
		<td bgcolor="<%= adminColor("pink") %>" width="100" >기본계약서</td>
	    <td bgcolor="#FFFFFF" colspan="4">
	    <% if (oDftContractList.FResultCount>0) then %>
	    <%= oDftContractList.FOneItem.FcontractName %>
	    /
	    <font color="<%= oDftContractList.FOneItem.GetContractStateColor %>" title="<%= oDftContractList.FOneItem.GetStateActiondate %>"><%= oDftContractList.FOneItem.GetContractStateName %></font>
	    /
	    <%= oDftContractList.FOneItem.FcontractDate %>
	    <% end if %>
	    </td>
	    <td bgcolor="#FFFFFF" align="right" width="80"></td>
	</tr>
	<tr >
		<td bgcolor="<%= adminColor("pink") %>" width="100" >온라인계약여부</td>
		<td bgcolor="#FFFFFF" colspan="4">
		<% if (onContractSheet.FResultCount>0) then %>
			<%= onContractSheet.FOneItem.FcontractName %>
			/<%= fnContractStateName(onContractSheet.FOneItem.FCtrState) %>
			/ <%= onContractSheet.FOneItem.FcontractDate %>
		<% end if %>
	    </td>
		<td bgcolor="#FFFFFF" align="right" width="80">
			<% if (onContractSheet.FResultCount>0) then %>
			<input type="button" class="button" value="계약서보기" onClick="RegContract('<%= designer %>','<%=groupid%>');">
			<% else %>
			<input type="button" class="button" value="계약서등록" onClick="RegContract('<%= designer %>','<%=groupid%>');">
			<% end if %>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("pink") %>">
		<td width="90"></td>
		<td >기본마진</td>
		<td colspan="3">추가 마진</td>
		<td>정산일</td>
	</tr>
	<tr>
		<td  bgcolor="<%= adminColor("pink") %>" >온라인<br>기본마진</td>
		<td align="center" bgcolor="#FFFFFF" >
			<% DrawBrandMWUCombo "maeipdiv",opartner.FOneItem.Fmaeipdiv %>
			<input type="text" class="text" name="defaultmargine" value="<%= opartner.FOneItem.Fdefaultmargine %>" size="4" style="text-align:right" onKeyUp="chgAddMargin(this)">%
		</td>

		<td width="60" align="center" bgcolor="#FFFFFF">매입<br><input type="text" class="text" name="M_margin" value="<%= opartner.FOneItem.FM_margin %>" size="4" style="text-align:right">%</td>
		<td width="60" align="center" bgcolor="#FFFFFF">위탁<br><input type="text" class="text" name="W_margin" value="<%= opartner.FOneItem.FW_margin %>" size="4" style="text-align:right">%</td>
		<td width="60" align="center" bgcolor="#FFFFFF">업체배송<br><input type="text" class="text" name="U_margin" value="<%= opartner.FOneItem.FU_margin %>" size="4" style="text-align:right">%</td>

	    <td align="center" rowspan="2" bgcolor="#FFFFFF" >익월 <%= opartner.FOneItem.Fjungsan_date %></td>
	</tr>
	<tr>
	    <td  bgcolor="<%= adminColor("pink") %>" >배송비조건</td>
		<td align="center" bgcolor="#FFFFFF" height=25 colspan="4">
		<% if (opartner.FOneItem.FdefaultdeliveryType="9") then %>
		업체조건 <%= FormatNumber(opartner.FOneItem.FDefaultFreeBeasongLimit,0) %>원 미만 <%= FormatNumber(opartner.FOneItem.FDefaultDeliverPay,0) %>
		<% elseif (opartner.FOneItem.FdefaultdeliveryType="7") then %>
		업체착불
		<% else %>
		기본정책 (텐배송 : 3만원 미만 2,000원 , 업체배송 : 무료)
		<% end if %>
		&nbsp;<input type="button" class="button" value="수정" onClick="PopBrandAdminUsingChange('<%= opartner.FOneItem.FID %>');">
		</td>
	</tr>
	<tr>
	    <td colspan=6" bgcolor="#FFFFFF"></td>
	</tr>
	<tr>
	    <td bgcolor="<%= adminColor("pink") %>" width="100" >오프라인계약여부</td>
	    <td bgcolor="#FFFFFF" colspan="4">
		<% if (offContractSheet.FResultCount>0) then %>
			<%= offContractSheet.FOneItem.FcontractName %>
			/<%= fnContractStateName(offContractSheet.FOneItem.FCtrState) %>
			/ <%= offContractSheet.FOneItem.FcontractDate %>
		<% end if %>
	    </td>
		<td bgcolor="#FFFFFF" align="right">
			<% if (offContractSheet.FResultCount>0) then %>
			<input type="button" class="button" value="계약서보기" onClick="RegContract('<%= designer %>','<%=groupid%>');">
			<% else %>
			<input type="button" class="button" value="계약서등록" onClick="RegContract('<%= designer %>','<%=groupid%>');">
			<% end if %>
		</td>
	</tr>
	<tr >
		<td bgcolor="<%= adminColor("pink") %>" >오프라인</td>
		<td bgcolor="#FFFFFF" colspan="4">
			<table border=0 cellspacing=0 cellpadding=0 class=a width=100%>
			<!--
			<tr>
				<td width="130"><b>직영점대표</b></td>
				<td width="130" align="center"><%= ooffontract.GetSpecialChargeDivName("streetshop000") %></td>
				<td align="center"><%= ooffontract.GetSpecialDefaultMargin("streetshop000") %> %</td>
			</tr>
			-->
			<% for i=0 to ooffontract.FResultCount-1 %>
			<tr>
				<td><%= ooffontract.FItemList(i).Fshopname %></td>
				<td align="center"><%= ooffontract.FItemList(i).GetChargeDivName() %></td>
				<td align="center"><%= ooffontract.FItemList(i).Fdefaultmargin %> %</td>
			</tr>
			<% next %>
			</table>
		</td>
		<td align="center" bgcolor="#FFFFFF">익월 <%= opartner.FOneItem.Fjungsan_date_off %></td>
	</tr>
	<tr>
		<td colspan="6" bgcolor="#FFFFFF" align="center"><input type="button" class="button" value="브랜드정보 저장" onclick="SaveBrandInfo(frmbrand);"></td>
	</tr>
	</form>
	</table>
<% end if %>

<br>

<% if (pcuserdiv="9999_14") then %>
	<table width="<%= itableWidth %>" align="center" cellpadding="1" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr bgcolor="#FFFFFF">
		<td height="25" colspan="4">**매입처(아카데미) 추가정보**</td>
		<input type="hidden" name="selltype" value="<%= opartner.FOneItem.Fselltype %>"> <!-- 자동설정 -->
		<input type="hidden" name="purchasetype" value="<%= opartner.FOneItem.Fpurchasetype %>">
		<input type="hidden" name="catecode" value="<%= opartner.FOneItem.Fcatecode %>"> <!-- 전시안함 -->
		<input type="hidden" name="offcatecode" value="<%= opartner.FOneItem.Foffcatecode %>"> <!-- 전시안함 -->
		<input type="hidden" name="offmduserid" value="<%= opartner.FOneItem.Foffmduserid %>"> <!-- OFFMD 없음 -->

		<input type="hidden" name="isextusing" value="<%= opartner.FOneItem.Fisextusing %>"> <!-- 제휴몰 사용안함 -->
		<input type="hidden" name="extstreetusing" value="<%= opartner.FOneItem.Fextstreetusing %>"> <!-- 제휴몰 Street 사용안함 -->
		<input type="hidden" name="specialbrand" value="<%= opartner.FOneItem.Fspecialbrand %>"> <!-- specialbrand 커뮤니티 사용안함 -->
		<input type="hidden" name="onlyflg" value="<%= opartner.FOneItem.Fonlyflg %>"> <!-- onlyflg 사용안함 -->
		<input type="hidden" name="artistflg" value="<%= opartner.FOneItem.Fartistflg %>"> <!-- artistflg 사용안함 -->
		<input type="hidden" name="kdesignflg" value="<%= opartner.FOneItem.Fkdesignflg %>"> <!-- kdesignflg 사용안함 -->

		<input type="hidden" name="maeipdiv" value="<%= opartner.FOneItem.Fmaeipdiv %>">         <!-- 매입구분 -->
		<input type="hidden" name="defaultmargine" value="<%= opartner.FOneItem.Fdefaultmargine %>">  <!-- 기본마진(유동) -->
		<input type="hidden" name="defaultdeliverytype" value="<%= opartner.FOneItem.Fdefaultdeliverytype %>">

		<input type="hidden" name="M_margin" value="<%= opartner.FOneItem.FM_margin %>">
		<input type="hidden" name="W_margin" value="<%= opartner.FOneItem.FW_margin %>">
		<input type="hidden" name="U_margin" value="<%= opartner.FOneItem.FU_margin %>">
	</tr>
	<tr >
		<td width="100" bgcolor="<%= adminColor("pink") %>"></td>
		<td bgcolor="#FFFFFF" ></td>
		<td width="100" bgcolor="<%= adminColor("pink") %>" >담당MD</td>
		<td bgcolor="#FFFFFF"><% drawSelectBoxCoWorker_OnOff "mduserid", opartner.FOneItem.Fmduserid, "fingers" %></td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("pink") %>">브랜드<br>사용여부</td>
		<td bgcolor="#FFFFFF">
		    <input type=radio name="isusing" value="Y" <%= CHKIIF(opartner.FOneItem.Fisusing="Y","checked","") %> >Y
		    <input type=radio name="isusing" value="N" <%= CHKIIF(opartner.FOneItem.Fisusing="N","checked","") %> >N</td>
		<td bgcolor="<%= adminColor("pink") %>">스트리트<br>표시여부<br>(브랜드운영관련)</td>
		<td bgcolor="#FFFFFF">
		    <input type=radio name="streetusing" value="Y" <%= CHKIIF(opartner.FOneItem.Fstreetusing="Y","checked","") %> >Y
		    <input type=radio name="streetusing" value="N" <%= CHKIIF(opartner.FOneItem.Fstreetusing="N","checked","") %> >N</td>
	</tr>
	<tr >
		<td width="120" bgcolor="#DDDDFF" rowspan="2">강좌 진행 여부</td>
		<td bgcolor="#FFFFFF" rowspan="2">
		<input type="radio" name="lec_yn" value="Y" <%= CHKIIF(opartner.FOneItem.Flec_yn="Y","checked","") %> onClick="clickLec(this)"> Y
		<input type="radio" name="lec_yn" value="N" <%= CHKIIF(opartner.FOneItem.Flec_yn="N","checked","") %> onClick="clickLec(this)"> N
		</td>
		<td width="120" bgcolor="#DDDDFF">강좌기본마진</td>
		<td bgcolor="#FFFFFF">
		<input type="text" name="lec_margin" value="<%= opartner.FOneItem.Flec_margin %>" size="4" maxlength="3"> (%)
		</td>
	</tr>
	<tr>
	    <td width="120" bgcolor="#DDDDFF">재료기본마진</td>
		<td bgcolor="#FFFFFF" width="200">
		<input type="text" name="mat_margin" value="<%= opartner.FOneItem.Fmat_margin %>" size="4" maxlength="3"> (%)
		</td>
	</tr>
	<tr >
		<td width="120" bgcolor="#DDDDFF" >DIY 진행 여부</td>
		<td  bgcolor="#FFFFFF" width="200" >
		<input type="radio" name="diy_yn" value="Y"  <%= CHKIIF(opartner.FOneItem.Fdiy_yn="Y","checked","") %> onClick="clickDiy(this);"> Y
		<input type="radio" name="diy_yn" value="N"  <%= CHKIIF(opartner.FOneItem.Fdiy_yn="N","checked","") %>  onClick="clickDiy(this);"> N
		</td>
		<td width="120" bgcolor="#DDDDFF">기본마진</td>
		<td bgcolor="#FFFFFF" width="200">
		<input type="text" name="diy_margin" value="<%= opartner.FOneItem.Fdiy_margin %>" size="4" maxlength="5"> (%) [부가세포함]
		</td>
	</tr>
	<tr id="iDiyDlv" style="display:<%= CHKIIF(opartner.FOneItem.Fdiy_yn="Y","","none") %> ">
		<td width="120" bgcolor="#DDDDFF">DIY배송구분</td>
		<td colspan="3" bgcolor="#FFFFFF">
		<select name="diy_dlv_gubun" onChange="chkCompdiygbn(this);">
		<option value="0" <%= CHKIIF(opartner.FOneItem.FdefaultDeliveryType="0","selected","") %>>기본(업체무료배송)
		<option value="9" <%= CHKIIF(opartner.FOneItem.FdefaultDeliveryType="9","selected","") %> >업체 조건배송
		</select>
		<br>
		<input type="hidden" name="pDFL" value="<%= opartner.FOneItem.FDefaultFreebeasongLimit %>">
		<input type="hidden" name="pDDP" value="<%= opartner.FOneItem.FdefaultDeliverPay %>">
		<input type="text" name="DefaultFreebeasongLimit" value="<%= opartner.FOneItem.FDefaultFreebeasongLimit %>" size="9" maxlength="9">원 이상 무료배송
		/미만 배송비 <input type="text" name="DefaultDeliverPay" value="<%= opartner.FOneItem.FdefaultDeliverPay %>" size="9" maxlength="9">원
		</td>
	</tr>
	<tr>
		<td colspan="4" bgcolor="#FFFFFF" align="center"><input type="button" class="button" value="브랜드정보 저장" onclick="SaveBrandInfo(frmbrand);"></td>
	</tr>
	</form>
	</table>
<% end if %>

<br>

<% if (pcuserdiv="501_21") or (pcuserdiv="503_21") or (pcuserdiv="900_21") or (pcuserdiv="999_50") then %>
	<table width="<%= itableWidth %>" align="center" cellpadding="1" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr bgcolor="#FFFFFF">
		<td height="25" colspan="6">**<%=getPartnerCommCodeName("pcuserdiv",pcuserdiv)%> 추가정보**</td>

		<input type="hidden" name="catecode" value="<%= opartner.FOneItem.Fcatecode %>"> <!-- 전시안함 -->
		<input type="hidden" name="offcatecode" value="<%= opartner.FOneItem.Foffcatecode %>"> <!-- 전시안함 -->

		<input type="hidden" name="isextusing" value="<%= opartner.FOneItem.Fisextusing %>"> <!-- 제휴몰 사용안함 -->
		<input type="hidden" name="streetusing" value="<%= opartner.FOneItem.Fstreetusing %>"> <!--  Street 사용안함 -->
		<input type="hidden" name="extstreetusing" value="<%= opartner.FOneItem.Fextstreetusing %>"> <!-- 제휴몰 Street 사용안함 -->
		<input type="hidden" name="specialbrand" value="<%= opartner.FOneItem.Fspecialbrand %>"> <!-- specialbrand 커뮤니티 사용안함 -->
		<input type="hidden" name="onlyflg" value="<%= opartner.FOneItem.Fonlyflg %>"> <!-- onlyflg 사용안함 -->
		<input type="hidden" name="artistflg" value="<%= opartner.FOneItem.Fartistflg %>"> <!-- artistflg 사용안함 -->
		<input type="hidden" name="kdesignflg" value="<%= opartner.FOneItem.Fkdesignflg %>"> <!-- kdesignflg 사용안함 -->

		<input type="hidden" name="maeipdiv" value="<%= opartner.FOneItem.Fmaeipdiv %>">         <!-- 매입구분 -->
		<input type="hidden" name="defaultmargine" value="<%= opartner.FOneItem.Fdefaultmargine %>">  <!-- 기본마진(유동) -->

		<input type="hidden" name="M_margin" value="<%= opartner.FOneItem.FM_margin %>">
		<input type="hidden" name="W_margin" value="<%= opartner.FOneItem.FW_margin %>">
		<input type="hidden" name="U_margin" value="<%= opartner.FOneItem.FU_margin %>">

	</tr>
	<tr >
		<td width="100" bgcolor="<%= adminColor("pink") %>">브랜드 사용여부</td>
		<td bgcolor="#FFFFFF" width="200">
		    <input type=radio name="isusing" value="Y" <%= CHKIIF(opartner.FOneItem.Fisusing="Y","checked","") %> >Y
		    <input type=radio name="isusing" value="N" <%= CHKIIF(opartner.FOneItem.Fisusing="N","checked","") %> >N
		</td>
		<td width="100" bgcolor="<%= adminColor("pink") %>" >담당자(영업)</td>
		<td bgcolor="#FFFFFF"><% drawSelectBoxCoWorker_OnOffUserdiv "mduserid", opartner.FOneItem.Fmduserid, "sell" ,pcuserdiv %></td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("pink") %>">기본 정산방식</td>
		<td bgcolor="#FFFFFF">
			<% if (C_ADMIN_AUTH or C_OFF_AUTH or C_MD_AUTH) then %>
				<% drawPartnerCommCodeBox true,"selljungsantype","purchasetype",opartner.FOneItem.FpurchaseType,"" %>
				<font color="blue"><b>*</b></font>
			<% else %>
				<%= getPartnerCommCodeName("selljungsantype", opartner.FOneItem.FpurchaseType) %>
				<input type="hidden" name="purchasetype" value="<%= opartner.FOneItem.FpurchaseType %>">
			<% end if %>
		</td>
		<td bgcolor="<%= adminColor("pink") %>">기본 매출계정</td>
		<td bgcolor="#FFFFFF">
			<% if (C_ADMIN_AUTH or C_OFF_AUTH or C_MD_AUTH) then %>
				<% drawPartnerCommCodeBox true,"sellacccd","selltype",opartner.FOneItem.Fselltype,"" %>
				<font color="blue"><b>*</b></font>
			<% else %>
				<%= getPartnerCommCodeName("sellacccd", opartner.FOneItem.Fselltype) %>
				<input type="hidden" name="selltype" value="<%= opartner.FOneItem.Fselltype %>">
			<% end if %>
		</td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("pink") %>">기본 매출부서</td>
		<td bgcolor="#FFFFFF">
			<% if (C_ADMIN_AUTH or C_OFF_AUTH or C_MD_AUTH) then %>
				<%= fndrawSaleBizSecCombo(true,"sellBizCd",opartner.FOneItem.FsellBizCd,"") %>
				<font color="blue"><b>*</b></font>
			<% else %>
				<%= fndrawSaleBizSecComboName(opartner.FOneItem.FsellBizCd) %>
				<input type="hidden" name="sellBizCd" value="<%= opartner.FOneItem.FsellBizCd %>">
			<% end if %>
		</td>
		<td bgcolor="<%= adminColor("pink") %>">수수료</td>
		<td bgcolor="#FFFFFF">
			<% if (C_ADMIN_AUTH or C_OFF_AUTH or C_MD_AUTH) then %>
				<input type="text" name="commission" value="<%= opartner.FOneItem.getCommissionPro %>" size="4">%
				<font color="blue"><b>*</b></font>
			<% else %>
				<%= opartner.FOneItem.getCommissionPro %>%
				<input type="hidden" name="commission" value="<%= opartner.FOneItem.getCommissionPro %>">
			<% end if %>
		</td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("pink") %>">계산서발행방식</td>
		<td bgcolor="#FFFFFF">
			<% if (C_ADMIN_AUTH or C_OFF_AUTH or C_MD_AUTH) then %>
				<% drawPartnerCommCodeBox true,"taxevaltype","taxevaltype",opartner.FOneItem.Ftaxevaltype,"" %>
				<font color="blue"><b>*</b></font>
			<% else %>
				<%= getPartnerCommCodeName("taxevaltype", opartner.FOneItem.Ftaxevaltype) %>
				<input type="hidden" name="taxevaltype" value="<%= opartner.FOneItem.Ftaxevaltype %>">
			<% end if %>
		</td>
		<td bgcolor="<%= adminColor("pink") %>">(기타매출)정산방법</td>
		<td bgcolor="#FFFFFF">
			<% if (C_ADMIN_AUTH or C_OFF_AUTH or C_MD_AUTH) then %>
				<% drawPartnerCommCodeBox true,"etcjungsantype","etcjungsantype",opartner.FOneItem.Fetcjungsantype,"" %>
				<font color="blue"><b>*</b></font>
			<% else %>
				<%= getPartnerCommCodeName("etcjungsantype", opartner.FOneItem.Fetcjungsantype) %>
				<input type="hidden" name="etcjungsantype" value="<%= opartner.FOneItem.Fetcjungsantype %>">
			<% end if %>
		</td>
	</tr>
	<!-- 2013/10/31 추가 -->
	<tr>
		<td bgcolor="<%= adminColor("pink") %>">3pl 연결계정</td>
		<td bgcolor="#FFFFFF" colspan="3">
			<% if (C_ADMIN_AUTH or C_OFF_AUTH or C_MD_AUTH) then %>
				<% CALL drawPartner3plCompany("tplcompanyid",opartner.FOneItem.Ftplcompanyid,"") %>
				&nbsp;(필요한 경우만 선택)
				<font color="blue"><b>*</b></font>
			<% else %>
				opartner.FOneItem.Ftplcompanyid
				<input type="hidden" name="tplcompanyid" value="<%= opartner.FOneItem.Ftplcompanyid %>">
			<% end if %>
		</td>
	</tr>

	<% if (pcuserdiv="999_50") then %>
		<tr bgcolor="#FFFFFF">
			<td height="25" colspan="6">**<%=getPartnerCommCodeName("pcuserdiv",pcuserdiv)%> 제휴사 입점 정보**</td>

		</tr>
		<tr>
			<td bgcolor="<%= adminColor("pink") %>">제휴형태</td>
			<td bgcolor="#FFFFFF">
				<% if (C_ADMIN_AUTH or C_OFF_AUTH or C_MD_AUTH) then %>
					<% drawPartnerCommCodeBox true,"mallSellType","pmallSellType",opartner.FOneItem.FpmallSellType,"" %>
					<font color="blue"><b>*</b></font>
				<% else %>
					<%= getPartnerCommCodeName("mallSellType", opartner.FOneItem.FpmallSellType) %>
					<input type="hidden" name="pmallSellType" value="<%= opartner.FOneItem.FpmallSellType %>">
				<% end if %>
			</td>
			<td bgcolor="<%= adminColor("pink") %>">연동방식</td>
			<td bgcolor="#FFFFFF">
				<% if (C_ADMIN_AUTH or C_OFF_AUTH or C_MD_AUTH) then %>
					<% drawPartnerCommCodeBox true,"pcomType","pcomType",opartner.FOneItem.FpcomType,"" %>
					<font color="blue"><b>*</b></font>
				<% else %>
					<%= getPartnerCommCodeName("pcomType", opartner.FOneItem.FpcomType) %>
					<input type="hidden" name="pcomType" value="<%= opartner.FOneItem.FpcomType %>">
				<% end if %>
			</td>
		</tr>
		<tr>
			<td bgcolor="<%= adminColor("pink") %>">제휴어드민URL</td>
			<td bgcolor="#FFFFFF" colspan="3">
			   <input type="text" name="padminUrl" value="<%= opartner.FOneItem.FpadminUrl %>" size="60" maxlength="120">
			</td>
		</tr>

		<tr>
			<td bgcolor="<%= adminColor("pink") %>">제휴어드민계정</td>
			<td bgcolor="#FFFFFF" colspan="">
			   ID <input type="text" name="padminId" value="<%= opartner.FOneItem.FpadminId %>" size="10" maxlength="32">
			   <% if (C_ADMIN_AUTH) then %><span style="padding-left:10px;">PW 사용안함</span><%END IF%>
			</td>
			<td bgcolor="<%= adminColor("pink") %>">주문처리담당</td>
			<td bgcolor="#FFFFFF">
		        <% drawSelectBoxCoWorker_OnOff "offmduserid", opartner.FOneItem.Foffmduserid, "sell" %>
			</td>
		</tr>
		<tr>
			<td bgcolor="<%= adminColor("pink") %>">기본 배송비 조건</td>
			<td bgcolor="#FFFFFF" colspan="3">
				<input type="hidden" name="defaultdeliverytype" value="9">
				<% if (C_ADMIN_AUTH or C_OFF_AUTH or C_MD_AUTH) then %>
					<input type="text" name="defaultFreeBeasongLimit" value="<%= opartner.FOneItem.FdefaultFreeBeasongLimit %>" size="8" maxlength="7">원 미만 구매시
					배송료 <input type="text" name="defaultDeliverPay" value="<%= opartner.FOneItem.FdefaultDeliverPay %>" size="7" maxlength="7">원
					<font color="blue"><b>*</b></font>
				<% else %>
					<%= opartner.FOneItem.FdefaultFreeBeasongLimit %> 원 미만 구매시 배송료 <%= opartner.FOneItem.FdefaultDeliverPay %> 원
					<input type="hidden" name="defaultFreeBeasongLimit" value="<%= opartner.FOneItem.FdefaultFreeBeasongLimit %>">
					<input type="hidden" name="defaultDeliverPay" value="<%= opartner.FOneItem.FdefaultDeliverPay %>">
				<% end if %>
			</td>
		</tr>
	<% end if %>

	<tr height="30">
		<td colspan="6" bgcolor="#FFFFFF" align="center"><input type="button" class="button" value="브랜드정보 저장" onclick="SaveBrandInfo(frmbrand);"></td>
	</tr>
	</form>
	</table>
<% end if %>

<% if (pcuserdiv="501_21") or (pcuserdiv="503_21") then %>
	<table width="<%= itableWidth %>" align="center" cellpadding="1" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr bgcolor="#FFFFFF">
		<td height="25" colspan="6">**매장 정보**</td>
	</tr>
	<tr>
	    <td height="25" colspan="6" align="right"  bgcolor="#FFFFFF">
		<input type="button" value="매장정보 보기 " onclick="popShopInfo('<%=designer%>');">
		</td>
	</tr>
	</table>
<% end if %>

<%
set opartner = Nothing
set ogroup = Nothing
set ooffontract = Nothing
set oDftContractList = Nothing
set onContractSheet = Nothing
set offContractSheet = Nothing
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->

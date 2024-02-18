<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 오프라인 매장 계약관리
' Hieditor : 2009.04.07 서동석 생성
'			 2011.01.21 한용민 수정
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshopchargecls.asp"-->
<!-- #include virtual="/lib/classes/partners/partnerusercls.asp"-->
<%
dim shopid, designer, mode, groupflag, page ,i, CProtoTypeShop, k ,shopdiv, shopusing, research
dim MainProtoTypeShopExists, defaultCenterMwDiv, MainProtoTypeShopComm_cd, offadminopen , itemregyn
dim Proto_comm_cd, Proto_defaultmargin, Proto_defaultsuplymargin, Proto_Etcjunsandetail
	page        = RequestCheckVar(request("page"),9)
	shopid      = RequestCheckVar(request("shopid"),32)
	designer    = RequestCheckVar(request("designer"),32)
	shopusing   = RequestCheckVar(request("shopusing"),1)
	research    = RequestCheckVar(request("research"),9)

if page="" then page=1
if (research="") then shopusing="Y"

dim oshopinfo, shopPurchaseType
set oshopinfo = new CPartnerUser
	oshopinfo.FRectDesignerID = shopid
	oshopinfo.GetOnePartnerNUser
	if oshopinfo.FresultCount > 0 then
		shopPurchaseType = oshopinfo.FOneItem.FpurchaseType
	end if

'//매장상세
dim ochargeuser
set ochargeuser = new COffShopChargeUser
	ochargeuser.FRectShopID     = shopid
	ochargeuser.FRectDesigner   = designer
	ochargeuser.GetOffShopDesignerList1

If ochargeuser.FTotalCount > 0 Then

	'//매장구분이 대표매장(MAIN) 인지 체크
	CProtoTypeShop = getoffshop_commoncodegroup("shopdiv", "MAIN", ochargeuser.FItemList(0).fshopdiv ,"")

	'//대표 매장일때만 shopdiv 를 넣어줌
	if CProtoTypeShop then
		shopdiv = ochargeuser.FItemList(0).fshopdiv
	end if
End If

'//대표마진 표시용
dim oOffshopdiv
set oOffshopdiv = new COffShopChargeUser
	oOffshopdiv.FPageSize        =   500
	oOffshopdiv.FRectDesigner    = designer
	oOffshopdiv.frectcodegroup = "MAIN"
	oOffshopdiv.Getoffshopdivmainlist

'//매장리스트
dim oOffMargin
set oOffMargin = new COffShopChargeUser
	oOffMargin.FPageSize        =   100
	oOffMargin.FRectDesigner    = designer
	oOffMargin.frectshopdiv = shopdiv
	oOffMargin.FRectShopusing  = shopusing
	oOffMargin.GetOffShopDesignerList2

'' 마진 변경 로그
dim oOffMarginLog
set oOffMarginLog = new COffShopChargeUser
	oOffMarginLog.FPageSize        = 10
	oOffMarginLog.FCurrPage        = page
	oOffMarginLog.FRectShopid      = shopid
	oOffMarginLog.FRectDesigner    = designer
	oOffMarginLog.GetOffShopMarginLogList

k=0

MainProtoTypeShopExists = false

dim prevMonthJungsanLogExist : prevMonthJungsanLogExist = False
prevMonthJungsanLogExist = IsExistPrevMonthShopJungsanLog(shopid, designer)

dim IsAdminAuth : IsAdminAuth = False
if C_ADMIN_AUTH then
	IsAdminAuth = True
end if

%>

<script type="text/javascript">

function editShopInfo(frm, ishopdiv){
    if (frm.comm_cd.value=="B011"){
        alert('사용 불가능한 정산구분입니다.(텐바이텐위탁)');
	    frm.comm_cd.focus();
	    return;
    }

    if ((frm.comm_cd.value!="B012")&&(frm.comm_cd.value!="B031")){
        if ((frm.orgcomm_cd.value=="")||(frm.orgcomm_cd.value!=frm.comm_cd.value)){
            //신규설정시 업체위탁 또는 출고매입만 가능함.
            <% if Not IsAdminAuth then %>
            alert('계약 신규설정시 업체위탁(B012) 또는 출고매입(B031) 만 가능합니다.');
            return;
			<% else %>
			if (confirm('관리자권한\n\n계약 신규설정시 업체위탁(B012) 또는 출고매입(B031) 만 가능합니다.\n계속 진행하시겠습니까?') != true) {
				return;
			}
            <% end if %>
        }
    }

    if ((frm.comm_cd.value=="B023")&&(<%= LCASE(C_ADMIN_AUTH) %>)){
        //가맹점 매입 (POS 개별상품)
        if (!confirm('가맹점 매입은 가맹점에서 개별 매입하는 경우만 사용 가능합니다. 계속하시겠습니까?')){
            return;
        }
    }else{
        if (frm.Proto_comm_cd.value.length<1 && <%= LCASE(not(C_ADMIN_AUTH)) %>){
            alert('대표샵 정산구분이 설정되지 않았습니다. 먼저 대표샵을 설정하신 후 사용하세요.');
    		return;
        }
    }

	if (frm.comm_cd.value.length<1){
		alert('업체 정산구분을 선택하세요.');
		frm.comm_cd.focus();
		return;
	}

    if (((ishopdiv=="5")||(ishopdiv=="6")||(ishopdiv=="7")||(ishopdiv=="8"))&&((frm.comm_cd.value!="B031")&&(frm.comm_cd.value!="B023"))){
        alert('도매나 해외 샵은 출고정산만 가능합니다.');
		frm.comm_cd.focus();
		return;
    }

    //가맹점은 텐위,업매 사용 불가
	<% if (session("ssBctId")<>"icommang") and (session("ssBctId")<>"coolhas") and (session("ssBctId")<>"gundolly") and (session("ssBctId")<>"mgrseul") then %>
	    if ((ishopdiv=="3")||(ishopdiv=="5")){
	        if ((frm.comm_cd.value=="B011")||(frm.comm_cd.value=="B022")){
	            alert('가맹점에 사용 불가능한 정산구분입니다.(텐바이텐위탁, 업체매입)');
			    frm.comm_cd.focus();
			    return;
	        }
	    }
	<% end if %>

	//정윤정 추가 2014.04.17
	if ( ishopdiv!="1" && ishopdiv!="2"){
		if((frm.jungsan_gubun.value == "간이과세") ||(frm.jungsan_gubun.value == "원천징수")){
			alert("정산구분이 [간이과세] 또는 [원천징수]인 브랜드ID는  직영점일때만 가능합니다.");
			return;
		}
	}

    //대표 마진과 다른지 여부 : 업체 위탁이 아닌경우 다를 수 없음 B013 :: 출고위탁
    if ((frm.comm_cd.value=="B012")||(frm.comm_cd.value=="B023")||(frm.comm_cd.value=="B013")){
        //업체위탁은 다를 수 있음. //가맹점 매입 // 출고위탁
    }else{
        if (frm.comm_cd.value!=frm.Proto_comm_cd.value && <%= LCASE(not(C_ADMIN_AUTH)) %>){
            alert('대표샵 정산 구분과 일치 하지 않습니다.1');
		    frm.comm_cd.focus();
		    return;
        }

        if (frm.defaultmargin.value*1!=frm.Proto_defaultmargin.value*1){
            <% if (session("ssBctId")<>"icommang") and (session("ssBctId")<>"coolhas") and (session("ssBctId")<>"gundolly") and (session("ssBctId")<>"tozzinet") and (session("ssBctId")<>"hrkang97") then %>
                //alert('대표샵 마진을과 일치 하지 않습니다.');
    		    //frm.defaultmargin.focus();
    		    //return;
		    <% else %>
		    if (!confirm('대표샵 마진을과 일치 하지 않습니다. 계속하시겠습니까?')){
		        frm.defaultmargin.focus();
		        return;
		    }
		    <% end if %>
        }

        if (frm.defaultsuplymargin.value*1!=frm.Proto_defaultsuplymargin.value*1){
            <% if (session("ssBctId")<>"icommang") and (session("ssBctId")<>"coolhas") and (session("ssBctId")<>"gundolly") then %>
                //alert('대표샵 마진을과 일치 하지 않습니다.');
    		    //frm.defaultsuplymargin.focus();
    		    //return;
    	    <% else %>
		    if (!confirm('대표샵 마진을과 일치 하지 않습니다. 계속하시겠습니까?')){
		        frm.defaultsuplymargin.focus();
		        return;
		    }
		    <% end if %>
        }
    }

	<% if (session("ssBctId")<>"icommang") and (session("ssBctId")<>"coolhas") and (session("ssBctId")<>"tozzinet") and (session("ssBctId")<>"hrkang97")  then %>
		if (frm.comm_cd.value=='B022'){
		    if (frm.Proto_comm_cd.value!='B022'){   //조건 추가 2011-11-29
    			alert('업체 매입은 대표 마진이 업체 매입인경우 만 사용 가능 합니다. 서이사님 또는 이이사님께 문의요망');
    			frm.comm_cd.focus();
    			return;
			}
		}

		if (frm.comm_cd.value=='B013'){
		    if (frm.Proto_comm_cd.value!='B013'){   //조건 추가 2011-11-29
    			alert('출고 위탁은 대표 마진이 출고 위탁인경우 만 사용 가능 합니다. 서이사님 또는 이이사님께 문의요망');
    			frm.comm_cd.focus();
    			return;
			}
		}
	<% end if %>

	//2014.04.14 정윤정 추가(브랜드id: 간이과세, 원천징수일때 업체정산구분 업체위탁, 출고위탁만 가능하도록)
	if((frm.jungsan_gubun.value == "간이과세") ||(frm.jungsan_gubun.value == "원천징수")){
		if((frm.comm_cd.value!='B012')&&(frm.comm_cd.value!='B013')){
		alert('브랜드ID의 정산구분이 [간이과세] 또는 [원천징수]인 경우\n업체정산구분은 [업체위탁]  또는 [출고위탁]만 가능합니다.');
		frm.comm_cd.focus();
		return;
		}
	}

	if ((frm.jungsan_gubun.value == "간이과세") && (frm.shopPurchaseType.value == "7")) {
		alert("=================================================\n\n에러!!! : 매장 정산방식이 출고가매출입니다.\n\n간이과세 브랜드는 등록불가입니다.\n\n=================================================");
		return;
	}

	<% if (session("ssBctId")<>"icommang") and (session("ssBctId")<>"coolhas") and (session("ssBctId")<>"tozzinet") and (session("ssBctId")<>"hrkang97") then %>
		if ((frm.onlinemwdiv.value=='M')&&((frm.comm_cd.value!='B031')&&(frm.comm_cd.value!='B012'))){
			alert('온라인 매입 인경우 매입출고 정산 또는 업체위탁만 가능합니다.');
			frm.chargediv.focus();
			return;
		}
	<% end if %>


	if ((frm.defaultmargin.value*1<0)||(frm.defaultmargin.value*1>100)){
		alert('기본 마진을 입력하세요.[0~100]');
		frm.defaultmargin.focus();
		return;
	}

	if ((frm.defaultsuplymargin.value*1<0)||(frm.defaultsuplymargin.value*1>100)){
		alert('기본 공급마진을 입력하세요.[0~100]');
		frm.defaultsuplymargin.focus();
		return;
	}

	// 역마진 띵소 제외
	<%
		If ochargeuser.FTotalCount>0 Then
			if (ochargeuser.FItemList(0).FDesignerId <> "ithinkso") then %>
    if (frm.defaultmargin.value*1<frm.defaultsuplymargin.value*1) {
	    alert("공급마진이 매입마진 보다 클 수 없습니다.");
		frm.defaultsuplymargin.focus();
		return;
	}
	<%		end if
		end if
	%>

	var ret = confirm('수정하시겠습니까?');
	if (ret){
		frm.submit();
	}
}

function delShopInfo(frm){
	frm.mode.value="del";
	var ret = confirm('삭제 하시겠습니까?');
	if (ret){
		frm.submit();
	}
}

function NextPage(page){
    document.frm.page.value = page;
    document.frm.submit();
}

function delShopInfoArr(frm){
    var selExists = false;
    var protoChecked = false;
    //대표샵을 선택한경우 나머지샵을 삭제 안할 수 없음.

    if (frm.cksel.length){
        protoChecked = (frm.cksel[0].checked);

        for (var i=0;i<frm.cksel.length;i++){
            if (frm.cksel[i].checked){
                selExists = true;

            }else{
                if ((protoChecked)&&(frm.precomm_cd[i].value.length>1)){
                    alert('대표샵을 삭제할 경우 계약된 전체샵을 삭제하셔야 합니다.');
                    frm.cksel[i].focus();
                    return;
                }
            }
        }

        if (!selExists){
            alert('선택 내역이 없습니다.');
            return;
        }
    }

    if (confirm('삭제 하시겠습니까?')){
        frm.mode.value="delArr";
        frm.submit();
    }
}

function editShopInfoArr(frm){
    var selExists = false;

    //대표샵 마진
    var oshopid, ocomm_cd, odefaultmargin, odefaultsuplymargin, oshopdiv;

    if (frm.cksel.length){
        if (frm.cksel[0].checked){
            oshopid         = frm.shopid[0].value;
            ocomm_cd        = frm.comm_cd[0].value;
            odefaultmargin  = frm.defaultmargin[0].value;
            odefaultsuplymargin = frm.defaultsuplymargin[0].value;
            oshopdiv        = frm.shopdiv[0].value;

            if (ocomm_cd=="B011"){
                alert('사용 불가능한 정산구분입니다.(텐바이텐위탁)');
    		    frm.comm_cd[0].focus();
    		    return;
            }

            //대표 마진 체크 : 직영만, 텐위, 업위 가능.
            //가맹점은 텐위,업매 사용 불가
			<% if (session("ssBctId")<>"icommang") and (session("ssBctId")<>"coolhas") and (session("ssBctId")<>"gundolly") then %>
			            if ((oshopdiv=="4")||(oshopdiv=="6")){
			                if ((ocomm_cd=="B011")||(ocomm_cd=="B022")){
			                    alert('가맹점에 사용 불가능한 정산구분입니다.(텐바이텐위탁, 업체매입)');
			        		    frm.comm_cd[0].focus();
			        		    return;
			                }
			            }
			<% end if %>

			//정윤정 추가 2014.04.17
			if ( oshopdiv!="1" && oshopdiv!="2"){
				if((frm.jungsan_gubun.value == "간이과세") ||(frm.jungsan_gubun.value == "원천징수")){
					alert("정산구분이 [간이과세] 또는 [원천징수]인 브랜드ID는  직영점일때만 가능합니다.1");
					return;
				}
			}
        }else{

            oshopid  = frm.shopid[0].value;
            ocomm_cd = frm.Proto_comm_cd.value;
            odefaultmargin = frm.Proto_defaultmargin.value;
            odefaultsuplymargin = frm.Proto_defaultsuplymargin.value;

            if (ocomm_cd=="B011"){
                alert('사용 불가능한 정산구분입니다.(텐바이텐위탁)');
    		    return;
            }
        }

		var shopdiv='<%= shopdiv %>';

        for (var i=0;i<frm.cksel.length;i++){
            if (frm.cksel[i].checked){
                selExists = true;

                if (frm.comm_cd[i].value.length<1){
                    alert('정산구분을 선택하세요.');
                    frm.comm_cd[i].focus();
                    return;
                }

	            if (frm.comm_cd[i].value=="B011"){
	                alert('사용 불가능한 정산구분입니다.(텐바이텐위탁)');
	                frm.comm_cd[i].focus();
	    		    return;
	            }

                if ((frm.comm_cd[i].value!="B012")&&(frm.comm_cd[i].value!="B031")){
                	//ygentshop 이 아닐경우에만
                	if (shopdiv!='14'){
	                    if ((frm.precomm_cd[i].value=="")||(frm.precomm_cd[i].value!=frm.comm_cd[i].value)){
	                        //신규설정시 업체위탁 또는 출고매입만 가능함. 201210 서동석 수정 - 사장님 요청
							<% if Not IsAdminAuth then %>
							alert('계약 신규설정시 업체위탁(B012) 또는 출고매입(B031) 만 가능합니다.');
							return;
							<% else %>
							if (confirm('관리자권한\n\n관리자 권한으로 진행 하시겠습니까?') != true) {
								return;
							}
							<% end if %>
						}
					}
                }

                if (((frm.shopdiv[i].value=="5")||(frm.shopdiv[i].value=="6")||(frm.shopdiv[i].value=="7")||(frm.shopdiv[i].value=="8"))&&(frm.comm_cd[i].value!="B031")){
                    alert('도매나 해외 샵은 출고정산만 가능합니다.' + frm.shopdiv[i].value);
            		frm.comm_cd[i].focus();
            		return;
                }

                //정윤정 추가 2014.04.17
				if ( frm.shopdiv[i].value!="1" && frm.shopdiv[i].value!="2"){
					if((frm.jungsan_gubun.value == "간이과세") ||(frm.jungsan_gubun.value == "원천징수")){
						alert("정산구분이 [간이과세] 또는 [원천징수]인 브랜드ID는  직영점일때만 가능합니다.");
						return;
					}
				}

               	//2014.04.14 정윤정 추가(브랜드id: 간이과세, 원천징수일때 업체정산구분 업체위탁, 출고위탁만 가능하도록)
				if((frm.jungsan_gubun.value == "간이과세") ||(frm.jungsan_gubun.value == "원천징수")){
					if((frm.comm_cd[i].value!='B012')&&(frm.comm_cd[i].value!='B013')){
						alert('브랜드ID의 정산구분이 [간이과세] 또는 [원천징수]인 경우\n업체정산구분은 [업체위탁]  또는 [출고위탁]만 가능합니다.');
						frm.comm_cd[i].focus();
						return;
					}
				}

				if ((frm.jungsan_gubun.value == "간이과세") && (frm.shopPurchaseType[i].value == "7")) {
					alert("=================================================\n\n에러!!! : 매장 정산방식이 출고가매출입니다.\n\n간이과세 브랜드는 등록불가입니다.\n\n=================================================");
					return;
				}

                //881동도트레이딩 출고마진 10%
                if ((frm.shopid[i].value=="streetshop881")&&(frm.defaultsuplymargin[i].value!="10")){
                    alert('동도트레이딩 출고마진 10% 만 가능');
            		frm.defaultsuplymargin[i].focus();
            		return;
                }

                if ((frm.defaultmargin[i].value*1<0)||(frm.defaultmargin[i].value*1>100)){
            		alert('기본 매입마진을 입력하세요.(0~100)');
            		frm.defaultmargin[i].focus();
            		return;
            	}

            	if ((frm.defaultsuplymargin[i].value*1<0)||(frm.defaultsuplymargin[i].value*1>100)){
            		alert('기본 공급마진을 입력하세요.(0~100)');
            		frm.defaultsuplymargin[i].focus();
            		return;
            	}

            	if (frm.defaultmargin[i].value*1<frm.defaultsuplymargin[i].value*1){
            	    alert('공급마진이 매입마진 보다 클 수 없습니다.');
            		frm.defaultsuplymargin[i].focus();
            		return;
            	}

				if (frm.chkdefaultbeasongdiv[i].checked == true) {
					// 업체배송
					frm.defaultbeasongdiv[i].value = "2";
				} else {
					frm.defaultbeasongdiv[i].value = "0";
				}

	            //대표 마진과 다른지 여부 : 업체 위탁이 아닌경우 다를 수 없음
                if (frm.comm_cd[i].value=="B012"){
                    //업체위탁은 다를 수 있음.
                }else{
	                    if (frm.comm_cd[i].value!=ocomm_cd){

	                        alert('대표샵 정산 구분과 일치 하지 않습니다.');

	                        <% if (session("ssBctId")="icommang") or (session("ssBctId")="coolhas") or (session("ssBctId")="tozzinet") or (session("ssBctId")="hrkang97") then %>
	                            if (confirm('관리자 권한으로 진행 하시겠습니까?')){
	                            }else{
	                                return;
	                            }

	                        <% else %>
		            		    frm.comm_cd[i].focus();
		            		    return;
		            		<% end if %>
	                    }

                    if (frm.defaultmargin[i].value*1!=odefaultmargin*1){
                        <% if (session("ssBctId")<>"icommang") and (session("ssBctId")<>"coolhas") and (session("ssBctId")<>"gundolly") then %>
                           // alert('대표샵 마진을과 일치 하지 않습니다.');
                		   // frm.defaultmargin[i].focus();
                		   // return;
            		    <% else %>
                            if (!confirm('대표샵 마진을과 일치 하지 않습니다. 계속하시겠습니까?')){
                		        frm.defaultsuplymargin[i].focus();
                		        return;
                		    }
            		    <% end if %>
                    }

                    if (frm.defaultsuplymargin[i].value*1!=odefaultsuplymargin*1){
                        <% if (session("ssBctId")<>"icommang") and (session("ssBctId")<>"coolhas") and (session("ssBctId")<>"gundolly") then %>
                            //alert('대표샵 마진을과 일치 하지 않습니다.');
                		    //frm.defaultsuplymargin[i].focus();
                		    //return;
                		<% else %>
                            if (!confirm('대표샵 마진을과 일치 하지 않습니다. 계속하시겠습니까?')){
                		        frm.defaultsuplymargin[i].focus();
                		        return;
                		    }
            		    <% end if %>
                    }
                }
            }
        }

        if (!selExists){
            alert('선택 내역이 없습니다.');
            return;
        }
    }else{
        alert('관리자 문의 요망');
    }

    if (confirm('저장 하시겠습니까?')){
        frm.mode.value="arredit";
        frm.submit();
    }
}

function AssignAllShop(){
    var oshopid, ocomm_cd, odefaultmargin, odefaultsuplymargin, ochkdefaultbeasongdiv;

	for (var i=0;i<frmeditArr.shopid.length;i++){
		if(frmeditArr.cksel[i].getAttribute("IsProtoType")=='Y') {
	        oshopid             	= frmeditArr.shopid[i].value;
	        ocomm_cd            	= frmeditArr.comm_cd[i].value;
	        odefaultmargin      	= frmeditArr.defaultmargin[i].value;
	        odefaultsuplymargin 	= frmeditArr.defaultsuplymargin[i].value;
	        ochkdefaultbeasongdiv	= frmeditArr.chkdefaultbeasongdiv[i].checked;
		}
	}

    if (ocomm_cd.length<1){
        alert('대표샵 정산구분을 선택하세요.');
        return;
    }
    if (odefaultmargin*1<1){
		alert('기본 매입마진을 입력하세요.');
		return;
	}
	if (odefaultsuplymargin*1<1){
		alert('기본 공급마진을 입력하세요.');
		return;
	}

	for (var i=0;i<frmeditArr.shopid.length;i++){
		if(frmeditArr.cksel[i].checked) {
		    if ((frmeditArr.shopisusing[i].value=="Y")&&(frmeditArr.shopid[i].value!=oshopid)){
		        if ((frmeditArr.comm_cd[i].value!="B012")&&(frmeditArr.comm_cd[i].value!="B022")){
	    	        document.all.trname[i].style.background='orange';
	    	        frmeditArr.cksel[i].checked = true;
	    	        frmeditArr.comm_cd[i].value = ocomm_cd;
	    	        frmeditArr.defaultmargin[i].value = odefaultmargin;
	    	        frmeditArr.defaultsuplymargin[i].value = odefaultsuplymargin;
	    	        frmeditArr.chkdefaultbeasongdiv[i].checked = ochkdefaultbeasongdiv;
		        }
		    }
		}
	}
}

function setDefaultShopMargin(frm){
    if (frm.Proto_comm_cd.value==""){
        alert('대표샵 마진이 설정되어 있지 않습니다.');
    }else{
        frm.comm_cd.value = frm.Proto_comm_cd.value;
        frm.defaultmargin.value = frm.Proto_defaultmargin.value;
        frm.defaultsuplymargin.value = frm.Proto_defaultsuplymargin.value;
    }
}

function ChangeDefaultCenterMwDiv(frm, MainProtoTypeShopComm_cd){
    //제약사항.
    if ((frm.defaultCenterMwDiv.checked)&&(MainProtoTypeShopComm_cd!="B031")&&(MainProtoTypeShopComm_cd!="B022")){
        alert('대표 정산구분이 출고분 정산/오프매입인 경우만 오프매입으로 사용 가능합니다.');
        return;
    }else{
        if (confirm('수정 하시겠습니까?')){
            frm.submit();
        }
    }
}

function ChangeOFFAdminOpen(frm, ShopComm_cd){
    if (ShopComm_cd.length<1){
        alert('대표 계약 조건 저장후 사용가능.');
        return;
    }

    if (confirm('수정 하시겠습니까?')){
        frm.submit();
    }
}

function totalCheck(){
	var f = document.frmeditArr;
	var objStr = "cksel";
	var chk_flag = true;
	for(var i=0; i<f.elements.length; i++) {
		if(f.elements[i].name == objStr) {
			if(!f.elements[i].checked) {
				chk_flag = f.elements[i].checked;
				break;
			}
		}
	}

	for(var i=0; i < f.elements.length; i++) {
		if(f.elements[i].name == objStr) {
			if(chk_flag) {
				f.elements[i].checked = false;
			} else {
				f.elements[i].checked = true;
			}
		}
	}
}

</script>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="research" value="on">
<input type="hidden" name="page">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		ShopID : <% Call NewDrawSelectBoxDesignerwithNameAndUserDIV("shopid",shopid, "21") %>
		&nbsp;
		브랜드ID : <% drawSelectBoxDesignerwithName "designer",designer  %>
		<br>Shop운영여부 : <% drawSelectBoxUsingYN "shopusing",shopusing %>
	</td>
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="frm.submit();">
	</td>
</tr>
</form>
</table>

<br>

<% if oOffMargin.FResultCount>0 then %>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		<b>* 기본계약조건</b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>구분</td>
    <td>대표계약조건</td>
    <td>매입마진(업체)</td>
    <td>출고마진(SHOP)</td>
    <td width="100"></td>
</tr>
<tr align="center" bgcolor="#FFFFFF">
    <td>온라인</td>
    <td><%= oOffMargin.FItemList(0).getMwName %></td>
    <td><%= oOffMargin.FItemList(0).Fonlinedefaultmargine %></td>
    <td></td>
    <td></td>
</tr>
<% if oOffshopdiv.FResultCount>0 then %>

<% for i=0 to oOffshopdiv.FResultCount -1 %>

	<%
	'//직영매장[대표] 일때만
	if getoffshopdiv(oOffshopdiv.FItemList(i).FShopID) ="2"  then
	    MainProtoTypeShopExists = true
	    MainProtoTypeShopComm_cd = oOffshopdiv.FItemList(i).FComm_cd
	    defaultCenterMwDiv = oOffshopdiv.FItemList(i).FdefaultCenterMwDiv
	    offadminopen = oOffshopdiv.FItemList(i).Fadminopen
	    itemregyn =  oOffshopdiv.FItemList(i).fitemregyn
	end if
	%>
	<tr align="center" bgcolor="#FFFFFF">
	    <td>
	    	<a href="?shopid=<%= oOffshopdiv.FItemList(i).FShopID %>&designer=<%= designer %>">
	    	<%= oOffshopdiv.FItemList(i).fshopdivname %></a>
	    </td>
	    <td><font color="<%= oOffshopdiv.FItemList(i).getChargeDivColor %>"><%= oOffshopdiv.FItemList(i).getChargeDivName %></font></td>
	    <td><%= oOffshopdiv.FItemList(i).FDefaultMargin %></td>
	    <td><%= oOffshopdiv.FItemList(i).FDefaultSuplyMargin %></td>
	    <td></td>
	</tr>
	<%
	'//해당매장이 대표매장 하위에 있는경우 기본마진 때려넣음
	if getoffshop_commoncodegroup("shopdiv", "SUB", oOffshopdiv.FItemList(i).fshopdiv ,chkIIF(ochargeuser.FTotalCount>0,ochargeuser.FItemList(0).fshopdiv,"")) then
	    Proto_comm_cd            = oOffshopdiv.FItemList(i).FComm_cd
	    Proto_defaultmargin      = oOffshopdiv.FItemList(i).FDefaultMargin
	    Proto_defaultsuplymargin = oOffshopdiv.FItemList(i).FDefaultSuplyMargin
	    Proto_Etcjunsandetail    = oOffshopdiv.FItemList(i).FEtcjunsandetail
	end if
	%>
<% next %>

<% end if %>

<% if (MainProtoTypeShopExists) then %>
<form name="frmAdmOpenEdit" method="post" action="popshopupcheinfo_process.asp">
<input type="hidden" name="mode" value="offadminopen">
<input type="hidden" name="shopid" value="streetshop000">
<input type="hidden" name="designer" value="<%= designer %>">
<tr align="center" bgcolor="#FFFFFF">
    <td>어드민 권한</td>
    <td colspan="4" align="left">
	    <input type="checkbox" value="Y" name="offadminopen" <%= ChkIIF(offadminopen="Y","checked","") %> >오프라인 어드민 오픈
	    <input type="checkbox" value="Y" name="itemregyn" <%= ChkIIF(itemregyn="Y","checked","") %> >오프라인 상품 등록 권한 부여
	    <br>
	    (업체위탁,텐바이텐위탁,매장매입은 기본적으로 오픈되며, 상품등록이 가능합니다.)
	    <input type="button" class="button" value="저장" onClick="ChangeOFFAdminOpen(frmAdmOpenEdit,'<%= MainProtoTypeShopComm_cd %>');">

    </td>
</tr>
</form>
<% end if %>

<% if (MainProtoTypeShopExists) then %>
<form name="frmCtEdit" method="post" action="popshopupcheinfo_process.asp">
<input type="hidden" name="mode" value="defaultCenterMwdivChange">
<input type="hidden" name="shopid" value="streetshop000">
<input type="hidden" name="designer" value="<%= designer %>">
<tr align="center" bgcolor="#FFFFFF">
    <td>센터 매입 구분</td>
    <td colspan="4" align="left">
    <input type="checkbox" value="M" name="defaultCenterMwDiv" <%= ChkIIF(defaultCenterMwDiv="M","checked","") %> >오프 매입 정산&nbsp;
    (체크시 업체배송 및 오프상품 매입 입고 정산 / 해제시 출고분 or 위탁정산) &nbsp;
    <input type="button" class="button" value="저장" onClick="ChangeDefaultCenterMwDiv(frmCtEdit,'<%= MainProtoTypeShopComm_cd %>');">
    <br>

    </td>
</tr>
</form>
<% end if %>
</table>
<% end if %>

<% if (CProtoTypeShop) then %>
	<br>
	<!-- 액션 시작 -->
	<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
	<tr>
		<td align="left">
		</td>
		<td align="right">
			<input type="button" class="button" value="선택대표계약일괄적용" onclick="AssignAllShop();">
		</td>
	</tr>
	</table>
	<!-- 액션 끝 -->
	<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>" border="0">
	<form name="frmeditArr" method="post" action="popshopupcheinfo_process.asp">
	<input type="hidden" name="mode" value="arredit">
	<input type="hidden" name="designer" value="<%= designer %>">
	<input type="hidden" name="Proto_comm_cd" value="<%= Proto_comm_cd %>">
	<input type="hidden" name="Proto_defaultmargin" value="<%= Proto_defaultmargin %>">
	<input type="hidden" name="Proto_defaultsuplymargin" value="<%= Proto_defaultsuplymargin %>">
	<input type="hidden" name="jungsan_gubun" value="<%= chkIIF(ochargeuser.FTotalCount>0,ochargeuser.FItemList(0).Fjungsan_gubun,"") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="9">
			<b>* 상세내역</b>
		</td>
	</tr>
	<tr bgcolor="FFFFFF" align="center">
	    <td width="30"><input type="checkbox" name="ckall" onclick="totalCheck()"></td>
	    <td>매장명</td>
	    <td>ShopID</td>
	    <td width="140">계약</td>
	    <td width="60">매입마진</td>
	    <td width="60">공급마진</td>
	    <td width="30">업체<br>배송</td>
	    <td width="30">비고</td>
	</tr>
	<% for i=0 to oOffMargin.FResultCount-1 %>
	<input type="hidden" name="shopid" value="<%= oOffMargin.FItemList(i).FShopID %>">
	<input type="hidden" name="shopisusing" value="<%= oOffMargin.FItemList(i).FShopIsUsing %>">
	<input type="hidden" name="shopdiv" value="<%= oOffMargin.FItemList(i).FShopDiv %>">
	<input type="hidden" name="precomm_cd" value="<%= oOffMargin.FItemList(i).Fcomm_cd %>">
	<input type="hidden" name="shopPurchaseType" value="<%= oOffMargin.FItemList(i).FpurchaseType %>">

	<tr align="center" <%= ChkIIF(oOffMargin.FItemList(i).FShopIsUsing="Y","bgcolor='FFFFFF'","bgcolor='DDDDDD'") %> id="trname">
	    <td>
			<input type="checkbox" name="cksel" value="<%= k %>" <%= ChkIIF(oOffMargin.FItemList(i).FShopID=shopid,"checked","") %> IsProtoType="<% if oOffMargin.FItemList(i).IsProtoTypeShop then %>Y<% end if %>">
		</td>
	    <td align="left"><a href="?shopid=<%= oOffMargin.FItemList(i).FShopID %>&designer=<%= designer %>"><%= oOffMargin.FItemList(i).FShopName %></a></td>
	    <td><%= oOffMargin.FItemList(i).FShopid %></td>
	    <td>
	        <!-- <font color="<%= oOffMargin.FItemList(i).getChargeDivColor %>"><%= oOffMargin.FItemList(i).getChargeDivName %></font> -->
	        <% drawSelectBoxOFFJungsanCommCD "comm_cd",oOffMargin.FItemList(i).Fcomm_cd %>
	    </td>
	    <td><input type="text" name="defaultmargin" class="text" value="<%= oOffMargin.FItemList(i).FDefaultMargin %>" size="4" maxlength="10"></td>
	    <td>
	    	<input type="hidden" name="defaultbeasongdiv" value="">
	    	<input type="text" name="defaultsuplymargin" class="text" value="<%= oOffMargin.FItemList(i).FDefaultSuplyMargin %>" size="4" maxlength="10">
	    </td>
	    <td><input type="checkbox" class="checkbox" name="chkdefaultbeasongdiv" class="text" value="2" <% if (oOffMargin.FItemList(i).Fdefaultbeasongdiv = "2") then %>checked<% end if %>></td>
	    <td></td>
	</tr>
	<% k = k + 1 %>

	<% next %>
	<tr bgcolor="<%= adminColor("tabletop") %>">
		<td align="center">기타<br>메모</td>
		<td bgcolor="#FFFFFF" colspan="7">
		<!-- 메모는 대표 코드에만 넣음!! -->
		<textarea name="etcjunsandetail" class="textarea" cols="90" rows="6"><%= Proto_Etcjunsandetail %></textarea>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td colspan="9" align="center">
			<input type="button" class="button" value="선택 샵 일괄 수정" onClick="editShopInfoArr(frmeditArr);">&nbsp;
			<input type="button" class="button" value="선택 샵 일괄 삭제" onClick="delShopInfoArr(frmeditArr);">
		</td>
	</tr>
	</form>
	</table>
<% end if %>

<% If ochargeuser.FTotalCount <> "0" Then %>
	<%
	if NOt (CProtoTypeShop) then
	'if NOt (CProtoTypeShop) or C_ADMIN_AUTH then		'//대표마진설정시에만
	%>
	<br>
	<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frmedit" method="post" action="popshopupcheinfo_process.asp">
	<input type="hidden" name="shopid" value="<%= shopid %>">
	<input type="hidden" name="designer" value="<%= designer %>">
	<input type="hidden" name="mode" value="edit">
	<input type="hidden" name="jungsan_gubun" value="<%= ochargeuser.FItemList(0).Fjungsan_gubun  %>">
	<input type="hidden" name="onlinemwdiv" value="<%= ochargeuser.FItemList(0).FOnlineMWDiv %>">
	<input type="hidden" name="Proto_comm_cd" value="<%= Proto_comm_cd %>">
	<input type="hidden" name="Proto_defaultmargin" value="<%= Proto_defaultmargin %>">
	<input type="hidden" name="Proto_defaultsuplymargin" value="<%= Proto_defaultsuplymargin %>">
	<input type="hidden" name="orgcomm_cd" value="<%= ochargeuser.FItemList(0).Fcomm_cd %>">
	<input type="hidden" name="shopPurchaseType" value="<%= shopPurchaseType %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15">
			<b>* 상세내역.</b>
		</td>
	</tr>

	<% if ochargeuser.FresultCount >0 then %>

	<tr bgcolor="<%= adminColor("tabletop") %>">
		<td width="100">ShopID</td>
		<td bgcolor="#FFFFFF"><%= ochargeuser.FItemList(0).Fshopid %></td>
	</tr>
	<tr bgcolor="<%= adminColor("tabletop") %>">
		<td width="100">브랜드ID</td>
		<td bgcolor="#FFFFFF"><%= ochargeuser.FItemList(0).FDesignerId %> (<%= ochargeuser.FItemList(0).FDesignerName %>)</td>
	</tr>
	<tr bgcolor="<%= adminColor("tabletop") %>" height="30">
		<td width="100">업체 정산 구분</td>
		<td bgcolor="#FFFFFF">
		    <% if (IsNULL(ochargeuser.FItemList(0).Fcomm_cd) or (ochargeuser.FItemList(0).Fcomm_cd="")) then %>
		    [마진 미설정]
		    <% else %>
		        <font color="<%= ochargeuser.FItemList(0).getJungsanDivColor %>"><%= ochargeuser.FItemList(0).getJungsanDivName %></font>
			<% end if %>
			<% if (prevMonthJungsanLogExist = True) and InStr(",icommang,coolhas,hrkang97,tozzinet,", ("," + session("ssBctId") + ",")) <= 0 then %>
				<input type="hidden" name="comm_cd" value="<%= ochargeuser.FItemList(0).Fcomm_cd %>">
				<font color="red">변경불가!</font> 강희란 팀장님에게 문의.
			<% else %>
		    	<% drawSelectBoxOFFJungsanCommCD "comm_cd",ochargeuser.FItemList(0).Fcomm_cd %>
				<input type="button" class="button" value="대표계약조건으로설정" onClick="setDefaultShopMargin(frmedit)">
				<% if (prevMonthJungsanLogExist = True) then %>
				[관리자뷰]
				<% end if %>
			<% end if %>
		</td>
	</tr>
	<tr bgcolor="<%= adminColor("tabletop") %>">
		<td width="100">매입마진(업체)</td>
		<td bgcolor="#FFFFFF">
			<input type="text" name="defaultmargin" class="text" value="<%= ochargeuser.FItemList(0).FDefaultMargin %>" size="4" maxlength="10" > %
			(업체로부터 공급받는 마진)
		</td>
	</tr>
	<tr bgcolor="<%= adminColor("tabletop") %>">
		<td width="100">출고마진(SHOP)</td>
		<td bgcolor="#FFFFFF">
			<input type="text" name="defaultsuplymargin" class="text" value="<%= ochargeuser.FItemList(0).FDefaultSuplyMargin %>" size="4" maxlength="10" > %
			(SHOP으로 공급하는 마진)
		</td>
	</tr>
	<tr bgcolor="<%= adminColor("tabletop") %>">
		<td width="100">배송구분(SHOP)</td>
		<td bgcolor="#FFFFFF">
			<% drawCheckBoxShopBeasongDiv "defaultbeasongdiv",ochargeuser.FItemList(0).Fdefaultbeasongdiv %>
		</td>
	</tr>
	<tr bgcolor="<%= adminColor("tabletop") %>">
		<td width="100">기타<br>메모</td>
		<td bgcolor="#FFFFFF">
		<textarea name="etcjunsandetail" class="textarea" cols="70" rows="6"><%= Proto_Etcjunsandetail %></textarea>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td colspan="2" align="center">
			<input type="button" class="button" value="수정" onClick="editShopInfo(frmedit,'<%= ochargeuser.FItemList(0).FShopDiv %>');">&nbsp;
			<input type="button" class="button" value="삭제" onClick="delShopInfo(frmedit);" <%= ChkIIf (IsNULL(ochargeuser.FItemList(0).Fcomm_cd) or (ochargeuser.FItemList(0).Fcomm_cd=""),"disabled","")  %>>
		</td>
	</tr>
	<% end if %>
	</form>
	</table>

	<% end if %>
	<!-- 변경 로그 -->
	<br>
	<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
    <tr height="25" bgcolor="FFFFFF">
		<td colspan="10">
			<b>* 마진변경 내역 </b>(계약 변경시 자동으로 기록됩니다.)
		</td>
	</tr>
    <tr bgcolor="FFFFFF" align="center">
        <td width="30">번호</td>
	    <td width="120">샵코드</td>
	    <td width="120">계약</td>
	    <td width="80">매입마진</td>
	    <td width="80">공급마진</td>
	    <td width="40">업체<br>배송</td>
	    <td width="40">Center<br>MW</td>
	    <td width="100">날짜</td>
	    <td width="100">등록자</td>
	    <td width="100">액션</td>
	</tr>
	<% for i=0 to oOffMarginLog.FResultCount-1 %>
	<tr bgcolor="FFFFFF" align="center">
	    <td ><%= oOffMarginLog.FItemList(i).FLogidx %></td>
	    <td ><%= oOffMarginLog.FItemList(i).FShopid %></td>
	    <td ><font color="<%= oOffMarginLog.FItemList(i).getJungsanDivColor %>"><%= oOffMarginLog.FItemList(i).getJungsanDivName %></font></td>
	    <td ><%= oOffMarginLog.FItemList(i).Fdefaultmargin %></td>
	    <td ><%= oOffMarginLog.FItemList(i).Fdefaultsuplymargin %></td>
	    <td ><% if (oOffMarginLog.FItemList(i).Fdefaultbeasongdiv = "2") then %>Y<% end if %></td>
	    <td ><%= oOffMarginLog.FItemList(i).FdefaultCenterMwdiv %></td>
	    <td ><%= Left(oOffMarginLog.FItemList(i).Fregdate,10) %></td>
	    <td ><%= oOffMarginLog.FItemList(i).Freguserid %></td>
	    <td ><%= oOffMarginLog.FItemList(i).getActFlagName %></td>
	</tr>
	<% next %>
	<tr bgcolor="#FFFFFF">
		<td colspan="13" align="center">
			<% if oOffMarginLog.HasPreScroll then %>
				<a href="javascript:NextPage('<%= oOffMarginLog.StarScrollPage-1 %>');">[pre]</a>
			<% else %>
				[pre]
			<% end if %>

			<% for i=0 + oOffMarginLog.StarScrollPage to oOffMarginLog.FScrollCount + oOffMarginLog.StarScrollPage - 1 %>
				<% if i>oOffMarginLog.FTotalpage then Exit for %>
				<% if CStr(page)=CStr(i) then %>
				<font color="red">[<%= i %>]</font>
				<% else %>
				<a href="javascript:NextPage('<%= i %>');">[<%= i %>]</a>
				<% end if %>
			<% next %>

			<% if oOffMarginLog.HasNextScroll then %>
				<a href="javascript:NextPage('<%= i %>');">[next]</a>
			<% else %>
				[next]
			<% end if %>
		</td>
	</tr>
	</table>

	<!--
	* 위탁샵의 위탁브랜드의 마진율을 설정하려고 할때, 000,800,870 의 대표마진이 없을경우,<br>
	streetshop000 / streetshop800 / streetshop870 화면을 띄워 대표마진 설정<br>
	관리자일 경우 이동할 수 있도록 하고, 관리자가 아닐경우, 관리자에게 요청하여 설정하라는 메세지를 띄운다<br>

	<p>

	<table width="100%" cellspacing="1" class="a" bgcolor=#3d3d3d>
	<tr bgcolor="#FFDDDD">
		<td>
			<table border=0 cellspacing=0 cellpadding=1 width=590 class=a>
			<tr>
				<td width="50%">* 정산구분 설정 </td>
				<td width="50%">* 업체로부터 공급받는 마진 설정 </td>
			</tr>
			<tr>
				<td align=center>
				<table border=0 cellspacing=1 cellpadding=1 bgcolor=#3d3d3d width=290 class=a>
				<tr bgcolor="#FFDDDD">
					<td width=100>기본(온라인)마진</td>
					<td width=100>정산구분</td>
					<td width=100>공급받는마진</td>
				</tr>
				<tr bgcolor="#FFDDDD">
					<td rowspan=2>매입</td>
					<td >텐바이텐매입</td>
					<td >기본마진과 <b>동일</b></td>
				</tr>
				<tr bgcolor="#FFDDDD">
					<td >업체위탁</td>
					<td >별도계약</td>
				</tr>
				<tr bgcolor="#FFDDDD">
					<td rowspan=2>위탁</td>
					<td >텐바이텐위탁<br></td>
					<td >기본마진 이상</td>
				</tr>
				<tr bgcolor="#FFDDDD">
					<td >업체위탁</td>
					<td >별도계약</td>
				</tr>
				</table>
				</td>
				<td valign=top>
				<table border=0 cellspacing=1 cellpadding=1 bgcolor=#3d3d3d width=290 class=a>
				<tr bgcolor="#FFDDDD">
					<td width=100>공급받는마진</td>
					<td width=100>샵 공급마진</td>
					<td width=100>비고</td>
				</tr>
				<tr bgcolor="#FFDDDD">
					<td >40% 이상</td>
					<td >35%</td>
					<td >조정가능(행사)</td>
				</tr>
				<tr bgcolor="#FFDDDD">
					<td >35%</td>
					<td >30%</td>
					<td ></td>
				</tr>
				<tr bgcolor="#FFDDDD">
					<td >30%</td>
					<td >25%</td>
					<td ></td>
				</tr>
				<tr bgcolor="#FFDDDD">
					<td >25%</td>
					<td >20%</td>
					<td ></td>
				</tr>
				</table>
				</td>
			</tr>
			<tr>
				<td colspan=2>* 오프라인대표(streetshop000) 및 가맹점대표(streetshop800)로 <b>마진 및 매입방식통일</b>예정(업체위탁별도)</td>
			</tr>
			</table>
		</td>
	</tr>
	</table>
	-->
<%
Else
	Response.Write "데이터가 없습니다."
End If

set oOffshopdiv = Nothing
set oOffMargin  = Nothing
set ochargeuser = Nothing
set oOffMarginLog = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->

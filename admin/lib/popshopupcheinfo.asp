<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : �������� ���� ������
' Hieditor : 2009.04.07 ������ ����
'			 2011.01.21 �ѿ�� ����
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

'//�����
dim ochargeuser
set ochargeuser = new COffShopChargeUser
	ochargeuser.FRectShopID     = shopid
	ochargeuser.FRectDesigner   = designer
	ochargeuser.GetOffShopDesignerList1

If ochargeuser.FTotalCount > 0 Then

	'//���屸���� ��ǥ����(MAIN) ���� üũ
	CProtoTypeShop = getoffshop_commoncodegroup("shopdiv", "MAIN", ochargeuser.FItemList(0).fshopdiv ,"")

	'//��ǥ �����϶��� shopdiv �� �־���
	if CProtoTypeShop then
		shopdiv = ochargeuser.FItemList(0).fshopdiv
	end if
End If

'//��ǥ���� ǥ�ÿ�
dim oOffshopdiv
set oOffshopdiv = new COffShopChargeUser
	oOffshopdiv.FPageSize        =   500
	oOffshopdiv.FRectDesigner    = designer
	oOffshopdiv.frectcodegroup = "MAIN"
	oOffshopdiv.Getoffshopdivmainlist

'//���帮��Ʈ
dim oOffMargin
set oOffMargin = new COffShopChargeUser
	oOffMargin.FPageSize        =   100
	oOffMargin.FRectDesigner    = designer
	oOffMargin.frectshopdiv = shopdiv
	oOffMargin.FRectShopusing  = shopusing
	oOffMargin.GetOffShopDesignerList2

'' ���� ���� �α�
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
        alert('��� �Ұ����� ���걸���Դϴ�.(�ٹ�������Ź)');
	    frm.comm_cd.focus();
	    return;
    }

    if ((frm.comm_cd.value!="B012")&&(frm.comm_cd.value!="B031")){
        if ((frm.orgcomm_cd.value=="")||(frm.orgcomm_cd.value!=frm.comm_cd.value)){
            //�űԼ����� ��ü��Ź �Ǵ� �����Ը� ������.
            <% if Not IsAdminAuth then %>
            alert('��� �űԼ����� ��ü��Ź(B012) �Ǵ� ������(B031) �� �����մϴ�.');
            return;
			<% else %>
			if (confirm('�����ڱ���\n\n��� �űԼ����� ��ü��Ź(B012) �Ǵ� ������(B031) �� �����մϴ�.\n��� �����Ͻðڽ��ϱ�?') != true) {
				return;
			}
            <% end if %>
        }
    }

    if ((frm.comm_cd.value=="B023")&&(<%= LCASE(C_ADMIN_AUTH) %>)){
        //������ ���� (POS ������ǰ)
        if (!confirm('������ ������ ���������� ���� �����ϴ� ��츸 ��� �����մϴ�. ����Ͻðڽ��ϱ�?')){
            return;
        }
    }else{
        if (frm.Proto_comm_cd.value.length<1 && <%= LCASE(not(C_ADMIN_AUTH)) %>){
            alert('��ǥ�� ���걸���� �������� �ʾҽ��ϴ�. ���� ��ǥ���� �����Ͻ� �� ����ϼ���.');
    		return;
        }
    }

	if (frm.comm_cd.value.length<1){
		alert('��ü ���걸���� �����ϼ���.');
		frm.comm_cd.focus();
		return;
	}

    if (((ishopdiv=="5")||(ishopdiv=="6")||(ishopdiv=="7")||(ishopdiv=="8"))&&((frm.comm_cd.value!="B031")&&(frm.comm_cd.value!="B023"))){
        alert('���ų� �ؿ� ���� ������길 �����մϴ�.');
		frm.comm_cd.focus();
		return;
    }

    //�������� ����,���� ��� �Ұ�
	<% if (session("ssBctId")<>"icommang") and (session("ssBctId")<>"coolhas") and (session("ssBctId")<>"gundolly") and (session("ssBctId")<>"mgrseul") then %>
	    if ((ishopdiv=="3")||(ishopdiv=="5")){
	        if ((frm.comm_cd.value=="B011")||(frm.comm_cd.value=="B022")){
	            alert('�������� ��� �Ұ����� ���걸���Դϴ�.(�ٹ�������Ź, ��ü����)');
			    frm.comm_cd.focus();
			    return;
	        }
	    }
	<% end if %>

	//������ �߰� 2014.04.17
	if ( ishopdiv!="1" && ishopdiv!="2"){
		if((frm.jungsan_gubun.value == "���̰���") ||(frm.jungsan_gubun.value == "��õ¡��")){
			alert("���걸���� [���̰���] �Ǵ� [��õ¡��]�� �귣��ID��  �������϶��� �����մϴ�.");
			return;
		}
	}

    //��ǥ ������ �ٸ��� ���� : ��ü ��Ź�� �ƴѰ�� �ٸ� �� ���� B013 :: �����Ź
    if ((frm.comm_cd.value=="B012")||(frm.comm_cd.value=="B023")||(frm.comm_cd.value=="B013")){
        //��ü��Ź�� �ٸ� �� ����. //������ ���� // �����Ź
    }else{
        if (frm.comm_cd.value!=frm.Proto_comm_cd.value && <%= LCASE(not(C_ADMIN_AUTH)) %>){
            alert('��ǥ�� ���� ���а� ��ġ ���� �ʽ��ϴ�.1');
		    frm.comm_cd.focus();
		    return;
        }

        if (frm.defaultmargin.value*1!=frm.Proto_defaultmargin.value*1){
            <% if (session("ssBctId")<>"icommang") and (session("ssBctId")<>"coolhas") and (session("ssBctId")<>"gundolly") and (session("ssBctId")<>"tozzinet") and (session("ssBctId")<>"hrkang97") then %>
                //alert('��ǥ�� �������� ��ġ ���� �ʽ��ϴ�.');
    		    //frm.defaultmargin.focus();
    		    //return;
		    <% else %>
		    if (!confirm('��ǥ�� �������� ��ġ ���� �ʽ��ϴ�. ����Ͻðڽ��ϱ�?')){
		        frm.defaultmargin.focus();
		        return;
		    }
		    <% end if %>
        }

        if (frm.defaultsuplymargin.value*1!=frm.Proto_defaultsuplymargin.value*1){
            <% if (session("ssBctId")<>"icommang") and (session("ssBctId")<>"coolhas") and (session("ssBctId")<>"gundolly") then %>
                //alert('��ǥ�� �������� ��ġ ���� �ʽ��ϴ�.');
    		    //frm.defaultsuplymargin.focus();
    		    //return;
    	    <% else %>
		    if (!confirm('��ǥ�� �������� ��ġ ���� �ʽ��ϴ�. ����Ͻðڽ��ϱ�?')){
		        frm.defaultsuplymargin.focus();
		        return;
		    }
		    <% end if %>
        }
    }

	<% if (session("ssBctId")<>"icommang") and (session("ssBctId")<>"coolhas") and (session("ssBctId")<>"tozzinet") and (session("ssBctId")<>"hrkang97")  then %>
		if (frm.comm_cd.value=='B022'){
		    if (frm.Proto_comm_cd.value!='B022'){   //���� �߰� 2011-11-29
    			alert('��ü ������ ��ǥ ������ ��ü �����ΰ�� �� ��� ���� �մϴ�. ���̻�� �Ǵ� ���̻�Բ� ���ǿ��');
    			frm.comm_cd.focus();
    			return;
			}
		}

		if (frm.comm_cd.value=='B013'){
		    if (frm.Proto_comm_cd.value!='B013'){   //���� �߰� 2011-11-29
    			alert('��� ��Ź�� ��ǥ ������ ��� ��Ź�ΰ�� �� ��� ���� �մϴ�. ���̻�� �Ǵ� ���̻�Բ� ���ǿ��');
    			frm.comm_cd.focus();
    			return;
			}
		}
	<% end if %>

	//2014.04.14 ������ �߰�(�귣��id: ���̰���, ��õ¡���϶� ��ü���걸�� ��ü��Ź, �����Ź�� �����ϵ���)
	if((frm.jungsan_gubun.value == "���̰���") ||(frm.jungsan_gubun.value == "��õ¡��")){
		if((frm.comm_cd.value!='B012')&&(frm.comm_cd.value!='B013')){
		alert('�귣��ID�� ���걸���� [���̰���] �Ǵ� [��õ¡��]�� ���\n��ü���걸���� [��ü��Ź]  �Ǵ� [�����Ź]�� �����մϴ�.');
		frm.comm_cd.focus();
		return;
		}
	}

	if ((frm.jungsan_gubun.value == "���̰���") && (frm.shopPurchaseType.value == "7")) {
		alert("=================================================\n\n����!!! : ���� �������� ��������Դϴ�.\n\n���̰��� �귣��� ��ϺҰ��Դϴ�.\n\n=================================================");
		return;
	}

	<% if (session("ssBctId")<>"icommang") and (session("ssBctId")<>"coolhas") and (session("ssBctId")<>"tozzinet") and (session("ssBctId")<>"hrkang97") then %>
		if ((frm.onlinemwdiv.value=='M')&&((frm.comm_cd.value!='B031')&&(frm.comm_cd.value!='B012'))){
			alert('�¶��� ���� �ΰ�� ������� ���� �Ǵ� ��ü��Ź�� �����մϴ�.');
			frm.chargediv.focus();
			return;
		}
	<% end if %>


	if ((frm.defaultmargin.value*1<0)||(frm.defaultmargin.value*1>100)){
		alert('�⺻ ������ �Է��ϼ���.[0~100]');
		frm.defaultmargin.focus();
		return;
	}

	if ((frm.defaultsuplymargin.value*1<0)||(frm.defaultsuplymargin.value*1>100)){
		alert('�⺻ ���޸����� �Է��ϼ���.[0~100]');
		frm.defaultsuplymargin.focus();
		return;
	}

	// ������ ��� ����
	<%
		If ochargeuser.FTotalCount>0 Then
			if (ochargeuser.FItemList(0).FDesignerId <> "ithinkso") then %>
    if (frm.defaultmargin.value*1<frm.defaultsuplymargin.value*1) {
	    alert("���޸����� ���Ը��� ���� Ŭ �� �����ϴ�.");
		frm.defaultsuplymargin.focus();
		return;
	}
	<%		end if
		end if
	%>

	var ret = confirm('�����Ͻðڽ��ϱ�?');
	if (ret){
		frm.submit();
	}
}

function delShopInfo(frm){
	frm.mode.value="del";
	var ret = confirm('���� �Ͻðڽ��ϱ�?');
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
    //��ǥ���� �����Ѱ�� ���������� ���� ���� �� ����.

    if (frm.cksel.length){
        protoChecked = (frm.cksel[0].checked);

        for (var i=0;i<frm.cksel.length;i++){
            if (frm.cksel[i].checked){
                selExists = true;

            }else{
                if ((protoChecked)&&(frm.precomm_cd[i].value.length>1)){
                    alert('��ǥ���� ������ ��� ���� ��ü���� �����ϼž� �մϴ�.');
                    frm.cksel[i].focus();
                    return;
                }
            }
        }

        if (!selExists){
            alert('���� ������ �����ϴ�.');
            return;
        }
    }

    if (confirm('���� �Ͻðڽ��ϱ�?')){
        frm.mode.value="delArr";
        frm.submit();
    }
}

function editShopInfoArr(frm){
    var selExists = false;

    //��ǥ�� ����
    var oshopid, ocomm_cd, odefaultmargin, odefaultsuplymargin, oshopdiv;

    if (frm.cksel.length){
        if (frm.cksel[0].checked){
            oshopid         = frm.shopid[0].value;
            ocomm_cd        = frm.comm_cd[0].value;
            odefaultmargin  = frm.defaultmargin[0].value;
            odefaultsuplymargin = frm.defaultsuplymargin[0].value;
            oshopdiv        = frm.shopdiv[0].value;

            if (ocomm_cd=="B011"){
                alert('��� �Ұ����� ���걸���Դϴ�.(�ٹ�������Ź)');
    		    frm.comm_cd[0].focus();
    		    return;
            }

            //��ǥ ���� üũ : ������, ����, ���� ����.
            //�������� ����,���� ��� �Ұ�
			<% if (session("ssBctId")<>"icommang") and (session("ssBctId")<>"coolhas") and (session("ssBctId")<>"gundolly") then %>
			            if ((oshopdiv=="4")||(oshopdiv=="6")){
			                if ((ocomm_cd=="B011")||(ocomm_cd=="B022")){
			                    alert('�������� ��� �Ұ����� ���걸���Դϴ�.(�ٹ�������Ź, ��ü����)');
			        		    frm.comm_cd[0].focus();
			        		    return;
			                }
			            }
			<% end if %>

			//������ �߰� 2014.04.17
			if ( oshopdiv!="1" && oshopdiv!="2"){
				if((frm.jungsan_gubun.value == "���̰���") ||(frm.jungsan_gubun.value == "��õ¡��")){
					alert("���걸���� [���̰���] �Ǵ� [��õ¡��]�� �귣��ID��  �������϶��� �����մϴ�.1");
					return;
				}
			}
        }else{

            oshopid  = frm.shopid[0].value;
            ocomm_cd = frm.Proto_comm_cd.value;
            odefaultmargin = frm.Proto_defaultmargin.value;
            odefaultsuplymargin = frm.Proto_defaultsuplymargin.value;

            if (ocomm_cd=="B011"){
                alert('��� �Ұ����� ���걸���Դϴ�.(�ٹ�������Ź)');
    		    return;
            }
        }

		var shopdiv='<%= shopdiv %>';

        for (var i=0;i<frm.cksel.length;i++){
            if (frm.cksel[i].checked){
                selExists = true;

                if (frm.comm_cd[i].value.length<1){
                    alert('���걸���� �����ϼ���.');
                    frm.comm_cd[i].focus();
                    return;
                }

	            if (frm.comm_cd[i].value=="B011"){
	                alert('��� �Ұ����� ���걸���Դϴ�.(�ٹ�������Ź)');
	                frm.comm_cd[i].focus();
	    		    return;
	            }

                if ((frm.comm_cd[i].value!="B012")&&(frm.comm_cd[i].value!="B031")){
                	//ygentshop �� �ƴҰ�쿡��
                	if (shopdiv!='14'){
	                    if ((frm.precomm_cd[i].value=="")||(frm.precomm_cd[i].value!=frm.comm_cd[i].value)){
	                        //�űԼ����� ��ü��Ź �Ǵ� �����Ը� ������. 201210 ������ ���� - ����� ��û
							<% if Not IsAdminAuth then %>
							alert('��� �űԼ����� ��ü��Ź(B012) �Ǵ� ������(B031) �� �����մϴ�.');
							return;
							<% else %>
							if (confirm('�����ڱ���\n\n������ �������� ���� �Ͻðڽ��ϱ�?') != true) {
								return;
							}
							<% end if %>
						}
					}
                }

                if (((frm.shopdiv[i].value=="5")||(frm.shopdiv[i].value=="6")||(frm.shopdiv[i].value=="7")||(frm.shopdiv[i].value=="8"))&&(frm.comm_cd[i].value!="B031")){
                    alert('���ų� �ؿ� ���� ������길 �����մϴ�.' + frm.shopdiv[i].value);
            		frm.comm_cd[i].focus();
            		return;
                }

                //������ �߰� 2014.04.17
				if ( frm.shopdiv[i].value!="1" && frm.shopdiv[i].value!="2"){
					if((frm.jungsan_gubun.value == "���̰���") ||(frm.jungsan_gubun.value == "��õ¡��")){
						alert("���걸���� [���̰���] �Ǵ� [��õ¡��]�� �귣��ID��  �������϶��� �����մϴ�.");
						return;
					}
				}

               	//2014.04.14 ������ �߰�(�귣��id: ���̰���, ��õ¡���϶� ��ü���걸�� ��ü��Ź, �����Ź�� �����ϵ���)
				if((frm.jungsan_gubun.value == "���̰���") ||(frm.jungsan_gubun.value == "��õ¡��")){
					if((frm.comm_cd[i].value!='B012')&&(frm.comm_cd[i].value!='B013')){
						alert('�귣��ID�� ���걸���� [���̰���] �Ǵ� [��õ¡��]�� ���\n��ü���걸���� [��ü��Ź]  �Ǵ� [�����Ź]�� �����մϴ�.');
						frm.comm_cd[i].focus();
						return;
					}
				}

				if ((frm.jungsan_gubun.value == "���̰���") && (frm.shopPurchaseType[i].value == "7")) {
					alert("=================================================\n\n����!!! : ���� �������� ��������Դϴ�.\n\n���̰��� �귣��� ��ϺҰ��Դϴ�.\n\n=================================================");
					return;
				}

                //881����Ʈ���̵� ����� 10%
                if ((frm.shopid[i].value=="streetshop881")&&(frm.defaultsuplymargin[i].value!="10")){
                    alert('����Ʈ���̵� ����� 10% �� ����');
            		frm.defaultsuplymargin[i].focus();
            		return;
                }

                if ((frm.defaultmargin[i].value*1<0)||(frm.defaultmargin[i].value*1>100)){
            		alert('�⺻ ���Ը����� �Է��ϼ���.(0~100)');
            		frm.defaultmargin[i].focus();
            		return;
            	}

            	if ((frm.defaultsuplymargin[i].value*1<0)||(frm.defaultsuplymargin[i].value*1>100)){
            		alert('�⺻ ���޸����� �Է��ϼ���.(0~100)');
            		frm.defaultsuplymargin[i].focus();
            		return;
            	}

            	if (frm.defaultmargin[i].value*1<frm.defaultsuplymargin[i].value*1){
            	    alert('���޸����� ���Ը��� ���� Ŭ �� �����ϴ�.');
            		frm.defaultsuplymargin[i].focus();
            		return;
            	}

				if (frm.chkdefaultbeasongdiv[i].checked == true) {
					// ��ü���
					frm.defaultbeasongdiv[i].value = "2";
				} else {
					frm.defaultbeasongdiv[i].value = "0";
				}

	            //��ǥ ������ �ٸ��� ���� : ��ü ��Ź�� �ƴѰ�� �ٸ� �� ����
                if (frm.comm_cd[i].value=="B012"){
                    //��ü��Ź�� �ٸ� �� ����.
                }else{
	                    if (frm.comm_cd[i].value!=ocomm_cd){

	                        alert('��ǥ�� ���� ���а� ��ġ ���� �ʽ��ϴ�.');

	                        <% if (session("ssBctId")="icommang") or (session("ssBctId")="coolhas") or (session("ssBctId")="tozzinet") or (session("ssBctId")="hrkang97") then %>
	                            if (confirm('������ �������� ���� �Ͻðڽ��ϱ�?')){
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
                           // alert('��ǥ�� �������� ��ġ ���� �ʽ��ϴ�.');
                		   // frm.defaultmargin[i].focus();
                		   // return;
            		    <% else %>
                            if (!confirm('��ǥ�� �������� ��ġ ���� �ʽ��ϴ�. ����Ͻðڽ��ϱ�?')){
                		        frm.defaultsuplymargin[i].focus();
                		        return;
                		    }
            		    <% end if %>
                    }

                    if (frm.defaultsuplymargin[i].value*1!=odefaultsuplymargin*1){
                        <% if (session("ssBctId")<>"icommang") and (session("ssBctId")<>"coolhas") and (session("ssBctId")<>"gundolly") then %>
                            //alert('��ǥ�� �������� ��ġ ���� �ʽ��ϴ�.');
                		    //frm.defaultsuplymargin[i].focus();
                		    //return;
                		<% else %>
                            if (!confirm('��ǥ�� �������� ��ġ ���� �ʽ��ϴ�. ����Ͻðڽ��ϱ�?')){
                		        frm.defaultsuplymargin[i].focus();
                		        return;
                		    }
            		    <% end if %>
                    }
                }
            }
        }

        if (!selExists){
            alert('���� ������ �����ϴ�.');
            return;
        }
    }else{
        alert('������ ���� ���');
    }

    if (confirm('���� �Ͻðڽ��ϱ�?')){
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
        alert('��ǥ�� ���걸���� �����ϼ���.');
        return;
    }
    if (odefaultmargin*1<1){
		alert('�⺻ ���Ը����� �Է��ϼ���.');
		return;
	}
	if (odefaultsuplymargin*1<1){
		alert('�⺻ ���޸����� �Է��ϼ���.');
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
        alert('��ǥ�� ������ �����Ǿ� ���� �ʽ��ϴ�.');
    }else{
        frm.comm_cd.value = frm.Proto_comm_cd.value;
        frm.defaultmargin.value = frm.Proto_defaultmargin.value;
        frm.defaultsuplymargin.value = frm.Proto_defaultsuplymargin.value;
    }
}

function ChangeDefaultCenterMwDiv(frm, MainProtoTypeShopComm_cd){
    //�������.
    if ((frm.defaultCenterMwDiv.checked)&&(MainProtoTypeShopComm_cd!="B031")&&(MainProtoTypeShopComm_cd!="B022")){
        alert('��ǥ ���걸���� ���� ����/���������� ��츸 ������������ ��� �����մϴ�.');
        return;
    }else{
        if (confirm('���� �Ͻðڽ��ϱ�?')){
            frm.submit();
        }
    }
}

function ChangeOFFAdminOpen(frm, ShopComm_cd){
    if (ShopComm_cd.length<1){
        alert('��ǥ ��� ���� ������ ��밡��.');
        return;
    }

    if (confirm('���� �Ͻðڽ��ϱ�?')){
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

<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="research" value="on">
<input type="hidden" name="page">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
		ShopID : <% Call NewDrawSelectBoxDesignerwithNameAndUserDIV("shopid",shopid, "21") %>
		&nbsp;
		�귣��ID : <% drawSelectBoxDesignerwithName "designer",designer  %>
		<br>Shop����� : <% drawSelectBoxUsingYN "shopusing",shopusing %>
	</td>
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="frm.submit();">
	</td>
</tr>
</form>
</table>

<br>

<% if oOffMargin.FResultCount>0 then %>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		<b>* �⺻�������</b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>����</td>
    <td>��ǥ�������</td>
    <td>���Ը���(��ü)</td>
    <td>�����(SHOP)</td>
    <td width="100"></td>
</tr>
<tr align="center" bgcolor="#FFFFFF">
    <td>�¶���</td>
    <td><%= oOffMargin.FItemList(0).getMwName %></td>
    <td><%= oOffMargin.FItemList(0).Fonlinedefaultmargine %></td>
    <td></td>
    <td></td>
</tr>
<% if oOffshopdiv.FResultCount>0 then %>

<% for i=0 to oOffshopdiv.FResultCount -1 %>

	<%
	'//��������[��ǥ] �϶���
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
	'//�ش������ ��ǥ���� ������ �ִ°�� �⺻���� ��������
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
    <td>���� ����</td>
    <td colspan="4" align="left">
	    <input type="checkbox" value="Y" name="offadminopen" <%= ChkIIF(offadminopen="Y","checked","") %> >�������� ���� ����
	    <input type="checkbox" value="Y" name="itemregyn" <%= ChkIIF(itemregyn="Y","checked","") %> >�������� ��ǰ ��� ���� �ο�
	    <br>
	    (��ü��Ź,�ٹ�������Ź,��������� �⺻������ ���µǸ�, ��ǰ����� �����մϴ�.)
	    <input type="button" class="button" value="����" onClick="ChangeOFFAdminOpen(frmAdmOpenEdit,'<%= MainProtoTypeShopComm_cd %>');">

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
    <td>���� ���� ����</td>
    <td colspan="4" align="left">
    <input type="checkbox" value="M" name="defaultCenterMwDiv" <%= ChkIIF(defaultCenterMwDiv="M","checked","") %> >���� ���� ����&nbsp;
    (üũ�� ��ü��� �� ������ǰ ���� �԰� ���� / ������ ���� or ��Ź����) &nbsp;
    <input type="button" class="button" value="����" onClick="ChangeDefaultCenterMwDiv(frmCtEdit,'<%= MainProtoTypeShopComm_cd %>');">
    <br>

    </td>
</tr>
</form>
<% end if %>
</table>
<% end if %>

<% if (CProtoTypeShop) then %>
	<br>
	<!-- �׼� ���� -->
	<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
	<tr>
		<td align="left">
		</td>
		<td align="right">
			<input type="button" class="button" value="���ô�ǥ����ϰ�����" onclick="AssignAllShop();">
		</td>
	</tr>
	</table>
	<!-- �׼� �� -->
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
			<b>* �󼼳���</b>
		</td>
	</tr>
	<tr bgcolor="FFFFFF" align="center">
	    <td width="30"><input type="checkbox" name="ckall" onclick="totalCheck()"></td>
	    <td>�����</td>
	    <td>ShopID</td>
	    <td width="140">���</td>
	    <td width="60">���Ը���</td>
	    <td width="60">���޸���</td>
	    <td width="30">��ü<br>���</td>
	    <td width="30">���</td>
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
		<td align="center">��Ÿ<br>�޸�</td>
		<td bgcolor="#FFFFFF" colspan="7">
		<!-- �޸�� ��ǥ �ڵ忡�� ����!! -->
		<textarea name="etcjunsandetail" class="textarea" cols="90" rows="6"><%= Proto_Etcjunsandetail %></textarea>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td colspan="9" align="center">
			<input type="button" class="button" value="���� �� �ϰ� ����" onClick="editShopInfoArr(frmeditArr);">&nbsp;
			<input type="button" class="button" value="���� �� �ϰ� ����" onClick="delShopInfoArr(frmeditArr);">
		</td>
	</tr>
	</form>
	</table>
<% end if %>

<% If ochargeuser.FTotalCount <> "0" Then %>
	<%
	if NOt (CProtoTypeShop) then
	'if NOt (CProtoTypeShop) or C_ADMIN_AUTH then		'//��ǥ���������ÿ���
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
			<b>* �󼼳���.</b>
		</td>
	</tr>

	<% if ochargeuser.FresultCount >0 then %>

	<tr bgcolor="<%= adminColor("tabletop") %>">
		<td width="100">ShopID</td>
		<td bgcolor="#FFFFFF"><%= ochargeuser.FItemList(0).Fshopid %></td>
	</tr>
	<tr bgcolor="<%= adminColor("tabletop") %>">
		<td width="100">�귣��ID</td>
		<td bgcolor="#FFFFFF"><%= ochargeuser.FItemList(0).FDesignerId %> (<%= ochargeuser.FItemList(0).FDesignerName %>)</td>
	</tr>
	<tr bgcolor="<%= adminColor("tabletop") %>" height="30">
		<td width="100">��ü ���� ����</td>
		<td bgcolor="#FFFFFF">
		    <% if (IsNULL(ochargeuser.FItemList(0).Fcomm_cd) or (ochargeuser.FItemList(0).Fcomm_cd="")) then %>
		    [���� �̼���]
		    <% else %>
		        <font color="<%= ochargeuser.FItemList(0).getJungsanDivColor %>"><%= ochargeuser.FItemList(0).getJungsanDivName %></font>
			<% end if %>
			<% if (prevMonthJungsanLogExist = True) and InStr(",icommang,coolhas,hrkang97,tozzinet,", ("," + session("ssBctId") + ",")) <= 0 then %>
				<input type="hidden" name="comm_cd" value="<%= ochargeuser.FItemList(0).Fcomm_cd %>">
				<font color="red">����Ұ�!</font> ����� ����Կ��� ����.
			<% else %>
		    	<% drawSelectBoxOFFJungsanCommCD "comm_cd",ochargeuser.FItemList(0).Fcomm_cd %>
				<input type="button" class="button" value="��ǥ����������μ���" onClick="setDefaultShopMargin(frmedit)">
				<% if (prevMonthJungsanLogExist = True) then %>
				[�����ں�]
				<% end if %>
			<% end if %>
		</td>
	</tr>
	<tr bgcolor="<%= adminColor("tabletop") %>">
		<td width="100">���Ը���(��ü)</td>
		<td bgcolor="#FFFFFF">
			<input type="text" name="defaultmargin" class="text" value="<%= ochargeuser.FItemList(0).FDefaultMargin %>" size="4" maxlength="10" > %
			(��ü�κ��� ���޹޴� ����)
		</td>
	</tr>
	<tr bgcolor="<%= adminColor("tabletop") %>">
		<td width="100">�����(SHOP)</td>
		<td bgcolor="#FFFFFF">
			<input type="text" name="defaultsuplymargin" class="text" value="<%= ochargeuser.FItemList(0).FDefaultSuplyMargin %>" size="4" maxlength="10" > %
			(SHOP���� �����ϴ� ����)
		</td>
	</tr>
	<tr bgcolor="<%= adminColor("tabletop") %>">
		<td width="100">��۱���(SHOP)</td>
		<td bgcolor="#FFFFFF">
			<% drawCheckBoxShopBeasongDiv "defaultbeasongdiv",ochargeuser.FItemList(0).Fdefaultbeasongdiv %>
		</td>
	</tr>
	<tr bgcolor="<%= adminColor("tabletop") %>">
		<td width="100">��Ÿ<br>�޸�</td>
		<td bgcolor="#FFFFFF">
		<textarea name="etcjunsandetail" class="textarea" cols="70" rows="6"><%= Proto_Etcjunsandetail %></textarea>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td colspan="2" align="center">
			<input type="button" class="button" value="����" onClick="editShopInfo(frmedit,'<%= ochargeuser.FItemList(0).FShopDiv %>');">&nbsp;
			<input type="button" class="button" value="����" onClick="delShopInfo(frmedit);" <%= ChkIIf (IsNULL(ochargeuser.FItemList(0).Fcomm_cd) or (ochargeuser.FItemList(0).Fcomm_cd=""),"disabled","")  %>>
		</td>
	</tr>
	<% end if %>
	</form>
	</table>

	<% end if %>
	<!-- ���� �α� -->
	<br>
	<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
    <tr height="25" bgcolor="FFFFFF">
		<td colspan="10">
			<b>* �������� ���� </b>(��� ����� �ڵ����� ��ϵ˴ϴ�.)
		</td>
	</tr>
    <tr bgcolor="FFFFFF" align="center">
        <td width="30">��ȣ</td>
	    <td width="120">���ڵ�</td>
	    <td width="120">���</td>
	    <td width="80">���Ը���</td>
	    <td width="80">���޸���</td>
	    <td width="40">��ü<br>���</td>
	    <td width="40">Center<br>MW</td>
	    <td width="100">��¥</td>
	    <td width="100">�����</td>
	    <td width="100">�׼�</td>
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
	* ��Ź���� ��Ź�귣���� �������� �����Ϸ��� �Ҷ�, 000,800,870 �� ��ǥ������ �������,<br>
	streetshop000 / streetshop800 / streetshop870 ȭ���� ��� ��ǥ���� ����<br>
	�������� ��� �̵��� �� �ֵ��� �ϰ�, �����ڰ� �ƴҰ��, �����ڿ��� ��û�Ͽ� �����϶�� �޼����� ����<br>

	<p>

	<table width="100%" cellspacing="1" class="a" bgcolor=#3d3d3d>
	<tr bgcolor="#FFDDDD">
		<td>
			<table border=0 cellspacing=0 cellpadding=1 width=590 class=a>
			<tr>
				<td width="50%">* ���걸�� ���� </td>
				<td width="50%">* ��ü�κ��� ���޹޴� ���� ���� </td>
			</tr>
			<tr>
				<td align=center>
				<table border=0 cellspacing=1 cellpadding=1 bgcolor=#3d3d3d width=290 class=a>
				<tr bgcolor="#FFDDDD">
					<td width=100>�⺻(�¶���)����</td>
					<td width=100>���걸��</td>
					<td width=100>���޹޴¸���</td>
				</tr>
				<tr bgcolor="#FFDDDD">
					<td rowspan=2>����</td>
					<td >�ٹ����ٸ���</td>
					<td >�⺻������ <b>����</b></td>
				</tr>
				<tr bgcolor="#FFDDDD">
					<td >��ü��Ź</td>
					<td >�������</td>
				</tr>
				<tr bgcolor="#FFDDDD">
					<td rowspan=2>��Ź</td>
					<td >�ٹ�������Ź<br></td>
					<td >�⺻���� �̻�</td>
				</tr>
				<tr bgcolor="#FFDDDD">
					<td >��ü��Ź</td>
					<td >�������</td>
				</tr>
				</table>
				</td>
				<td valign=top>
				<table border=0 cellspacing=1 cellpadding=1 bgcolor=#3d3d3d width=290 class=a>
				<tr bgcolor="#FFDDDD">
					<td width=100>���޹޴¸���</td>
					<td width=100>�� ���޸���</td>
					<td width=100>���</td>
				</tr>
				<tr bgcolor="#FFDDDD">
					<td >40% �̻�</td>
					<td >35%</td>
					<td >��������(���)</td>
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
				<td colspan=2>* �������δ�ǥ(streetshop000) �� ��������ǥ(streetshop800)�� <b>���� �� ���Թ������</b>����(��ü��Ź����)</td>
			</tr>
			</table>
		</td>
	</tr>
	</table>
	-->
<%
Else
	Response.Write "�����Ͱ� �����ϴ�."
End If

set oOffshopdiv = Nothing
set oOffMargin  = Nothing
set ochargeuser = Nothing
set oOffMarginLog = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->

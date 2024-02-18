<%@ language=vbscript %>
<% option explicit %>
<%
'###############################################
' PageName : event_register.asp
' Discription : I��(������) �̺�Ʈ ��� ���� ������
' History : 2019.01.22 ������
'			2019-10-02 ������	���ø� �������� ����
'###############################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/admin/eventmanage/event/v5/lib/admineventhead.asp"-->
<!-- #include virtual="/lib/classes/event/eventManageCls_V5.asp"-->
<!-- #include virtual="/admin/eventmanage/common/event_function_v5.asp"-->
<!-- #include virtual="/admin/eventmanage/event/v5/lib/jqueryui_include.asp"-->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenDepartmentCls.asp"-->
<%
dim cEvtCont, cEvtMultiCont, ArrcMultiContentsMenu, ix
dim eCode, ekind, elevel, isWeb, isMobile, isApp, emdid, emdnm, eDateView
dim esale, egift, ecoupon, eonlyten, eOneplusone, eFreedelivery, eBookingsell
dim eDiary, eNew, ecomment, ebbs, eitemps, eisblogurl, ename, eusing, eman
dim eSdate, eEdate, ePdate, stepdiv, estate, efwd, fullefwd, estimateSalePrice
dim blnReqPublish, edgid, edgstat1, edgstat2, edgid2, epsid, LinkURL, LinkURL2, LinkURL3
dim epsnm, edpnm, edpid, sWorkTag, edgnm, edgnm2, eCCNm, maxDepth, EvtCopyCode
dim salePer, saleCPer, eEtcitemimg, ebimgMo2014, ebrand, etag, DispCate, nocate, ImgCopyUser
dim enamesub, subcopyK, etemp, etemp_mo, mdbntype, mdbntypemo, mdtheme, mdthememo, kakaoTitle
dim themecolor, themecolormo, title_pc, title_mo, eSalePer, DispCateName, evt_mainIMG, marketing_event_kind
dim cEGroup, vYear, arrGroup, arrGroup_mo, togglediv, viewset, evt_type, CommentTitle, ImgCopyCode, multiconOpen, menudiv, menuidx

eCode = requestCheckVar(Request("eC"),10)
ekind = requestCheckVar(Request("eK"),10)
togglediv = requestCheckVar(Request("togglediv"),2) '�޴� ��ġ�� ������
viewset = requestCheckVar(Request("viewset"),1) '�̸����� ������
multiconOpen = requestCheckVar(Request("multiconOpen"),1) '��Ƽ������ ����
menudiv = requestCheckVar(Request("menudiv"),2) '��Ƽ������ �޴� ����
menuidx = requestCheckVar(Request("menuidx"),10) '��Ƽ������ �޴� �ε���

if viewset="" then viewset="M"
if togglediv="" then  togglediv="1"
if emdid = "" then
    emdid = session("ssBctId")
    emdnm = session("ssBctCname")
end if

maxDepth=2
esale = False
egift= False
ecoupon= False
eonlyten= False
eOneplusone= False
eFreedelivery= False
eBookingsell= False
eDiary= False
eNew= False
ecomment = False
ebbs    = False
eitemps = False
eisblogurl = False

IF eCode <> "" THEN
    set cEvtCont = new ClsEvent
    cEvtCont.FECode = eCode '�̺�Ʈ �ڵ�
    '�̺�Ʈ ���� ��������
    cEvtCont.fnGetEventCont
    ekind = cEvtCont.FEKind
    eman = cEvtCont.FEManager
    ename = db2html(cEvtCont.FEName) ' �̺�Ʈ��
    eusing = cEvtCont.FEUsing
    elevel = cEvtCont.FELevel
    isWeb = cEvtCont.FIsWeb
    isMobile = cEvtCont.FIsMobile
    isApp = cEvtCont.FIsApp
    eSdate = cEvtCont.FESDay
    eEdate = cEvtCont.FEEDay
    ePdate = cEvtCont.FEPDay
    estate = cEvtCont.FEState
    enamesub = db2html(cEvtCont.FENamesub) '�̺�Ʈ Ÿ��Ʋ ����ī�� �����
    subcopyK =  db2html(cEvtCont.FsubcopyK) '����ī�� �ѱ� PC

    IF datediff("d",now,eEdate) <0 THEN estate = 9 '�Ⱓ �ʰ��� ����ǥ��
    
    if ekind = 19 then
        isWeb = False
        isMobile = True
        isApp = True
        ekind = 1
    elseif ekind = 25 then
        isWeb = False
        isMobile = False
        isApp = True
        ekind = 1
    elseif ekind = 26 then
        isWeb = False
        isMobile = True
        isApp = False
        ekind = 1
    elseif not (isWeb  or  isMobile  or isApp) or (isNull(isWeb) and isNull(isMobile) and isNull(isApp))  then 
        isWeb = True
        isMobile = False
        isApp = False    
        ekind = 1
    end if        

    '�̺�Ʈ ȭ�鼳�� ���� ��������
    cEvtCont.fnGetEventDisplay
    '�⺻ ����
    esale       =   cEvtCont.FESale
    egift       =   cEvtCont.FEGift
    ecoupon     =   cEvtCont.FECoupon
    ecomment    =   cEvtCont.FECommnet
    ebbs        =   cEvtCont.FEBbs
    eitemps     =   cEvtCont.FEItemps
    eonlyten    = cEvtCont.FSisOnlyTen
    eDiary      = cEvtCont.FSisDiary
    eNew            = cEvtCont.FSisNew
    eisblogurl  = cEvtCont.FSisGetBlogURL
    eOneplusone     = cEvtCont.FEOneplusOne
    eFreedelivery   = cEvtCont.FEFreedelivery
    eBookingsell    = cEvtCont.FEBookingsell
    eDateView       = cEvtCont.FEdateview
    salePer = cEvtCont.FsalePer
    saleCPer = cEvtCont.FsaleCPer
    '����ڿ�
    blnReqPublish = cEvtCont.FisReqPublish
    emdid = cEvtCont.FEMdId
    emdnm = cEvtCont.FEMdName
    edgid = cEvtCont.FEDgId
    edgid2 = cEvtCont.FEDgId2
    edgstat1 = cEvtCont.FEDgStat1
    edgstat2 = cEvtCont.FEDgStat2
    epsid = cEvtCont.FEPsId
    edpid = cEvtCont.FEDpId
    sWorkTag = cEvtCont.FWorkTag
    edgnm = cEvtCont.FEDgName
    edgnm2 = cEvtCont.FEDgName2
    epsnm = cEvtCont.FEPsName
    edpnm = cEvtCont.FEDpName
    eCCNm = cEvtCont.FECCName
    efwd = nl2br(db2html(chrbyte(cEvtCont.FEFwd,100,"Y")))
    '����Ʈ�� ���
    eEtcitemimg = cEvtCont.FEtcitemimg
    ebimgMo2014 = cEvtCont.FEBImgMoListBanner
    '��ȹ�� ����
    ebrand = cEvtCont.FEBrand
    etag = db2html(cEvtCont.FETag)
    DispCate = cEvtCont.FEDispCate
    '��� Ÿ��Ʋ
    etemp = cEvtCont.FETemp
    etemp_mo = cEvtCont.FETemp_mo
    mdbntype = cEvtCont.Fmdbntype
    mdbntypemo = cEvtCont.Fmdbntypemo
    themecolor = cEvtCont.Fthemecolor
    themecolormo = cEvtCont.Fthemecolormo
    evt_mainIMG = cEvtCont.FEMImg
    mdtheme = cEvtCont.Fmdtheme
    mdthememo = cEvtCont.Fmdthememo
    marketing_event_kind = cEvtCont.Fmarketing_event_kind

    '���� ��� �̺�Ʈ �׸� ����
    cEvtCont.fnGetEventMDThemeInfo
    '��ȹ�� ����
    nocate = cEvtCont.Fnocate
    'SNS ����
    kakaoTitle = cEvtCont.Fkakao_title
    '��� Ÿ��Ʋ
    title_pc = cEvtCont.Ftitle_pc
    title_mo = cEvtCont.Ftitle_mo
    DispCateName = cEvtCont.getDispCategory(DispCate)
    estimateSalePrice = cEvtCont.FestimateSalePrice
    evt_type = cEvtCont.Feventtype_pc          '���Ľ����̼� �����ڷ� ���
    CommentTitle = cEvtCont.FECommentTitle    '���Ľ����̼� �����ü�� ���
    EvtCopyCode = cEvtCont.FEvtCopyCode     '������ �̺�Ʈ �ڵ�
    ImgCopyUser = cEvtCont.FEvtImgCopyUserid '�̹��� ���� ����
    ImgCopyCode = cEvtCont.FEvtImgCopyCode  '�̹��� ���� ��� �̺�Ʈ�ڵ�
    set cEvtCont = nothing

    if etemp="10" or etemp="6" then
    else
        if mdtheme<>"5" then
            response.redirect "/admin/eventmanage/event/v4/event_modify.asp?eC="&eCode
            response.end
        end if
    end if

    if etemp_mo="11" or etemp_mo="6" or etemp_mo="10" then
    else
        if mdthememo<>"5" then
            response.redirect "/admin/eventmanage/event/v4/event_modify.asp?eC="&eCode
            response.end
        end if
    end if

    set cEGroup = new ClsEventGroup
    cEGroup.FECode = eCode
    cEGroup.FEChannel = "P"
    arrGroup = cEGroup.fnGetEventItemGroup
    cEGroup.FEChannel = "M"
    arrGroup_mo = cEGroup.fnGetEventItemGroup
    vYear = cEGroup.FRegdate
    set cEGroup = nothing

    if (ekind = 1 or ekind = 23) and (eSale or ecoupon) then
        dim tmpename
        tmpename = Split(ename,"|") 
                 
        if Ubound(tmpename)>0 then
            ename = tmpename(0)
            eSalePer = tmpename(1)
         end if

    end if

    if stepdiv = "" then  stepdiv=1

    set cEvtMultiCont = new ClsMultiContentsMenu
    cEvtMultiCont.FRectEvtCode = eCode
    ArrcMultiContentsMenu=cEvtMultiCont.fnGetMultiContentsMenuList
    set cEvtMultiCont = nothing
end if 

if application("Svr_Info")="Dev" then
    LinkURL = "https://testm.10x10.co.kr"
    LinkURL2 = "https://2015www.10x10.co.kr"
    LinkURL3 = "http://testm.10x10.co.kr"
else
    LinkURL = "https://m.10x10.co.kr"
    LinkURL2 = "https://www.10x10.co.kr"
    LinkURL3 = "http://m.10x10.co.kr"
end if

public Function fnEventBarColorCode(themecolor)
    If themecolor="1" Then
        fnEventBarColorCode = "background-color:#ed6c6c;"
    ElseIf themecolor="2" Then
        fnEventBarColorCode = "background-color:#f385af;"
    ElseIf themecolor="3" Then
        fnEventBarColorCode = "background-color:#f3a056;"
    ElseIf themecolor="4" Then
        fnEventBarColorCode = "background-color:#e7b93c;"
    ElseIf themecolor="5" Then
        fnEventBarColorCode = "background-color:#8eba4a;"
    ElseIf themecolor="6" Then
        fnEventBarColorCode = "background-color:#43a251;"
    ElseIf themecolor="7" Then
        fnEventBarColorCode = "background-color:#50bdd1;"
    ElseIf themecolor="8" Then
        fnEventBarColorCode = "background-color:#5aa5ea;"
    ElseIf themecolor="9" Then
        fnEventBarColorCode = "background-color:#2672bf;"
    ElseIf themecolor="10" Then
        fnEventBarColorCode = "background-color:#2c5a85;"
    ElseIf themecolor="11" Then
        fnEventBarColorCode = "background-color:#848484;"
    ElseIf themecolor="12" Then
        fnEventBarColorCode = "background-image:url(http://webadmin.10x10.co.kr/images/event/img_grd_1.jpg);"
    ElseIf themecolor="13" Then
        fnEventBarColorCode = "background-image:url(http://webadmin.10x10.co.kr/images/event/img_grd_2.jpg);"
    ElseIf themecolor="14" Then
        fnEventBarColorCode = "background-image:url(http://webadmin.10x10.co.kr/images/event/img_grd_3.jpg);"
    ElseIf themecolor="15" Then
        fnEventBarColorCode = "background-image:url(http://webadmin.10x10.co.kr/images/event/img_grd_4.jpg);"
    ElseIf themecolor="16" Then
        fnEventBarColorCode = "background-image:url(http://webadmin.10x10.co.kr/images/event/img_grd_5.jpg);"
    ElseIf themecolor="17" Then
        fnEventBarColorCode = "background-image:url(http://webadmin.10x10.co.kr/images/event/img_grd_6.jpg);"
    ElseIf themecolor="18" Then
        fnEventBarColorCode = "background-image:url(http://webadmin.10x10.co.kr/images/event/img_grd_7.jpg);"
    ElseIf themecolor="19" Then
        fnEventBarColorCode = "background-image:url(http://webadmin.10x10.co.kr/images/event/img_grd_8.jpg);"
    ElseIf themecolor="20" Then
        fnEventBarColorCode = "background-image:url(http://webadmin.10x10.co.kr/images/event/img_grd_9.jpg);"
    ElseIf themecolor="21" Then
        fnEventBarColorCode = "background-image:url(http://webadmin.10x10.co.kr/images/event/img_grd_10.jpg);"
    ElseIf themecolor="22" Then
        fnEventBarColorCode = "background-image:url(http://webadmin.10x10.co.kr/images/event/img_grd_11.jpg);"
    ElseIf themecolor="23" Then
        fnEventBarColorCode = "background-image:url(http://webadmin.10x10.co.kr/images/event/img_grd_12.jpg);"
    ElseIf themecolor="24" Then
        fnEventBarColorCode = "background-image:url(http://webadmin.10x10.co.kr/images/event/img_grd_13.jpg);"
    ElseIf themecolor="25" Then
        fnEventBarColorCode = "background-image:url(http://webadmin.10x10.co.kr/images/event/img_grd_14.jpg);"
    ElseIf themecolor="26" Then
        fnEventBarColorCode = "background-image:url(http://webadmin.10x10.co.kr/images/event/img_grd_15.jpg);"
    ElseIf themecolor="27" Then
        fnEventBarColorCode = "background-image:url(http://webadmin.10x10.co.kr/images/event/img_grd_16.jpg);"
    Else
        fnEventBarColorCode = "background-color:#848484;"
    End If
End Function

function GetMenuDivName(menudiv)
    if menudiv="1" then
        GetMenuDivName="�̹��� �����̵�"
    elseif menudiv="2" then
        GetMenuDivName="����"
    elseif menudiv="3" then
        GetMenuDivName="�귣�彺�丮"
    elseif menudiv="4" then
        GetMenuDivName="��õ����Ʈ"
    elseif menudiv="5" then
        GetMenuDivName="�����Ϳ���"
    elseif menudiv="6" then
        GetMenuDivName="��� �����̵�"
    elseif menudiv="7" then
        GetMenuDivName="�̹��� & HTML"
    elseif menudiv="8" then
        GetMenuDivName="�̹��� ���ø� �����̵�"
    elseif menudiv="9" then
        GetMenuDivName="������"
    elseif menudiv="10" then
        GetMenuDivName="��������"
    elseif menudiv="11" then
        GetMenuDivName="�̹�����ũ"
    elseif menudiv="12" then
        GetMenuDivName="��ǰ ���� ����"
    elseif menudiv="13" then
        GetMenuDivName="�ǹ�"
    end if
end function
%>
<style type="text/css">
body {background-color:#ffffff;}
</style>
<script>
function fnBasicInfoSet(eventcode){
    var winbasicView;
    winbasicView = window.open('/admin/eventmanage/event/v5/popup/pop_event_basicinfo.asp?eC='+eventcode,'basicinfo','width=1024,height=830,scrollbars=yes,resizable=yes');
    winbasicView.focus();
}

function fnWorkerInfoSet(eventcode){
    if(eventcode==""){
        alert("�⺻���� ������ �����ؾ� ���� �����մϴ�.");
    }
    else{
        var winworkerView;
        winworkerView = window.open('/admin/eventmanage/event/v5/popup/pop_event_workerinfo.asp?eC='+eventcode,'workerinfo','width=1024,height=800,scrollbars=yes,resizable=yes');
        winworkerView.focus();
    }
}

function fnViewWorkMemo(){
    var winworkMemoView;
    winworkMemoView = window.open('/admin/eventmanage/event/v5/popup/pop_workMemo.asp?eC=<%=eCode%>','workmemowin','width=1024,height=600,scrollbars=yes,resizable=yes');
    winworkMemoView.focus();
}

function fnListBannerSet(eventcode){
    if(eventcode==""){
        alert("�⺻���� ������ �����ؾ� ���� �����մϴ�.");
    }
    else{
        var winListBannerView;
        winListBannerView = window.open('/admin/eventmanage/event/v5/popup/pop_event_listbanner.asp?eC='+eventcode,'bannerinfo','width=1024,height=650,scrollbars=yes,resizable=yes');
        winListBannerView.focus();
    }
}

function fnEventInfoSet(eventcode){
    if(eventcode==""){
        alert("�⺻���� ������ �����ؾ� ���� �����մϴ�.");
    }
    else{
        var winSettingView;
        winSettingView = window.open('/admin/eventmanage/event/v5/popup/pop_event_setting.asp?eC='+eventcode,'eventsetinfo','width=1024,height=650,scrollbars=yes,resizable=yes');
        winSettingView.focus();
    }
}

function fnEventSnsShareSet(eventcode) {
    const winSettingView = window.open('/admin/eventmanage/event/v5/popup/pop_event_share_sns.asp?eC='+eventcode,'eventsharesns','width=1024,height=650,scrollbars=yes,resizable=yes');
    winSettingView.focus();
}

function fnEventTopInfoSet(eventcode){
    if(eventcode==""){
        alert("�⺻���� ������ �����ؾ� ���� �����մϴ�.");
    }
    else{
        var winTopSettingView;
        <% if instr(ename,"�ָ�Ư��")>0 or instr(ename,"����Ư��")>0 then %>
        winTopSettingView = window.open('/admin/eventmanage/event/v5/popup/_pop_event_top_info.asp?eC='+eventcode+'&wm=M','eventtopinfo','width=1024,height=900,scrollbars=yes,resizable=yes');
        <% elseif instr(ename,"����Ż�")>0 then %>
        winTopSettingView = window.open('/admin/eventmanage/event/v5/popup/_pop_event_top_info.asp?eC='+eventcode+'&wm=M','eventtopinfo','width=1024,height=900,scrollbars=yes,resizable=yes');
        <% else %>
        winTopSettingView = window.open('/admin/eventmanage/event/v5/popup/pop_event_top_info.asp?eC='+eventcode+'&wm=M','eventtopinfo','width=1024,height=900,scrollbars=yes,resizable=yes');
        <% end if %>
        winTopSettingView.focus();
    }
}

function fnGiftInfoSet(eventcode){
    if(eventcode==""){
        alert("�⺻���� ������ �����ؾ� ���� �����մϴ�.");
    }
    else{
        var winGiftView;
        winGiftView = window.open('/admin/shopmaster/gift/giftList.asp?eC='+eventcode,'giftinfo','width=1000,height=600,scrollbars=yes,resizable=yes');
        winGiftView.focus();
    }
}

function fnContentsMenuSet(eventcode){
    if(eventcode==""){
        alert("�⺻���� ������ �����ؾ� ���� �����մϴ�.");
    }
    else{
        var winMenuView;
        winMenuView = window.open('/admin/eventmanage/event/v5/popup/pop_contentsmenu_setting.asp?eC='+eventcode,'menuinfo','width=1024,height=600,scrollbars=yes,resizable=yes');
        winMenuView.focus();
    }
}

function fnMultiContentsSet(eventcode,menuidx,menudiv){
    var winMultiContentsView;
    if(menudiv=="1"){
        <% if ekind="5" then %>
        winMultiContentsView = window.open('/admin/eventmanage/event/v5/popup/pop_rolling_image_culture.asp?eC='+eventcode+'&menuidx='+menuidx,'swifeimginfo','width=1024,height=500,scrollbars=yes,resizable=yes');
        <% else %>
        winMultiContentsView = window.open('/admin/eventmanage/event/v5/popup/pop_rolling_image.asp?eC='+eventcode+'&menuidx='+menuidx,'swifeimginfo','width=1024,height=500,scrollbars=yes,resizable=yes');
        <% end if %>
    }
    else if(menudiv=="2"){
        winMultiContentsView = window.open('/admin/eventmanage/event/v5/popup/pop_video.asp?eC='+eventcode+'&menuidx='+menuidx,'brandinfo','width=1024,height=530,scrollbars=yes,resizable=yes');
    }
    else if(menudiv=="3"){
        winMultiContentsView = window.open('/admin/eventmanage/event/v5/popup/pop_event_brandstoryinfo.asp?eC='+eventcode+'&menuidx='+menuidx,'brandinfo','width=1024,height=530,scrollbars=yes,resizable=yes');
    }
    else if(menudiv=="4"){
        winMultiContentsView = window.open('/admin/eventmanage/event/v5/popup/pop_event_grouptemplateinfo.asp?eC='+eventcode+'&menuidx='+menuidx,'grouptemplateinfo','width=1024,height=800,scrollbars=yes,resizable=yes');
    }
    else if(menudiv=="5"){
        winMultiContentsView = window.open('/admin/eventmanage/event/v5/popup/pop_event_customboxinfo.asp?eC='+eventcode+'&menuidx='+menuidx,'customboxinfo','width=1024,height=800,scrollbars=yes,resizable=yes');
    }
    else if(menudiv=="6"){
        winMultiContentsView = window.open('/admin/eventmanage/event/v5/template/slide/pop_mobile_top_slide.asp?eC='+eventcode+'&menuidx='+menuidx,'topslide','width=1450,height=800,scrollbars=yes,resizable=yes');
    }
    else if(menudiv=="7"){
        winMultiContentsView = window.open('/admin/eventmanage/event/v5/popup/pop_event_img_text_template.asp?eC='+eventcode+'&menuidx='+menuidx,'imgtxt','width=1024,height=600,scrollbars=yes,resizable=yes');
    }
    else if(menudiv=="8"){
        winMultiContentsView = window.open('/admin/eventmanage/event/v5/template/slide/pop_mobile_slide.asp?eC='+eventcode+'&menuidx='+menuidx,'imgtxt','width=1450,height=800,scrollbars=yes,resizable=yes');
    }
    else if(menudiv=="9"){
        winMultiContentsView = window.open('/admin/eventmanage/event/v5/template/addbanner/pop_mobile_addbanner.asp?eC='+eventcode+'&menuidx='+menuidx,'imgtxt','width=1450,height=800,scrollbars=yes,resizable=yes');
    }
    else if(menudiv=="10"){
        winMultiContentsView = window.open('/admin/eventmanage/event/v5/popup/pop_event_execfile.asp?eC='+eventcode+'&menuidx='+menuidx,'customboxinfo','width=1024,height=500,scrollbars=yes,resizable=yes');
    }
    else if(menudiv=="11"){
        winMultiContentsView = window.open('/admin/eventmanage/event/v5/popup/pop_event_imagelink.asp?eC='+eventcode+'&menuidx='+menuidx,'customboxinfo','width=1024,height=500,scrollbars=yes,resizable=yes');
    }
    winMultiContentsView.focus();
}

function fnMultiContentsDeviceSet(eventcode,menuidx,menudiv,device){
    var winMultiContentsView;
    if(menudiv=="1"){
        <% if ekind="5" then %>
        winMultiContentsView = window.open('/admin/eventmanage/event/v5/popup/pop_rolling_image_culture.asp?eC='+eventcode+'&menuidx='+menuidx,'swifeimginfo','width=1024,height=500,scrollbars=yes,resizable=yes');
        <% else %>
        winMultiContentsView = window.open('/admin/eventmanage/event/v5/popup/pop_rolling_image.asp?device='+device+'&eC='+eventcode+'&menuidx='+menuidx,'swifeimginfo','width=1024,height=500,scrollbars=yes,resizable=yes');
        <% end if %>
    }
    else if(menudiv=="2"){
        winMultiContentsView = window.open('/admin/eventmanage/event/v5/popup/pop_video.asp?eC='+eventcode+'&menuidx='+menuidx,'brandinfo','width=1024,height=530,scrollbars=yes,resizable=yes');
    }
    else if(menudiv=="3"){
        winMultiContentsView = window.open('/admin/eventmanage/event/v5/popup/pop_event_brandstoryinfo.asp?eC='+eventcode+'&menuidx='+menuidx,'brandinfo','width=1024,height=530,scrollbars=yes,resizable=yes');
    }
    else if(menudiv=="4"){
        winMultiContentsView = window.open('/admin/eventmanage/event/v5/popup/pop_event_grouptemplateinfo.asp?eC='+eventcode+'&menuidx='+menuidx,'grouptemplateinfo','width=1024,height=800,scrollbars=yes,resizable=yes');
    }
    else if(menudiv=="5"){
        winMultiContentsView = window.open('/admin/eventmanage/event/v5/popup/pop_event_customboxinfo.asp?eC='+eventcode+'&menuidx='+menuidx,'customboxinfo','width=1024,height=800,scrollbars=yes,resizable=yes');
    }
    else if(menudiv=="6"){
        if(device=="M"){
            winMultiContentsView = window.open('/admin/eventmanage/event/v5/template/slide/pop_mobile_top_slide.asp?eC='+eventcode+'&menuidx='+menuidx,'topslide','width=1450,height=800,scrollbars=yes,resizable=yes');
        }
        else{
            winMultiContentsView = window.open('/admin/eventmanage/event/v5/template/slide/pop_pcweb_top_slide.asp?eC='+eventcode+'&menuidx='+menuidx,'topslide','width=1450,height=800,scrollbars=yes,resizable=yes');
        }
    }
    else if(menudiv=="7"){
        winMultiContentsView = window.open('/admin/eventmanage/event/v5/popup/pop_event_img_text_template.asp?wm='+device+'&eC='+eventcode+'&menuidx='+menuidx,'imgtxt','width=1024,height=600,scrollbars=yes,resizable=yes');
    }
    else if(menudiv=="8"){
        if(device=="M"){
            winMultiContentsView = window.open('/admin/eventmanage/event/v5/template/slide/pop_mobile_slide.asp?eC='+eventcode+'&menuidx='+menuidx,'imgtxt','width=1450,height=800,scrollbars=yes,resizable=yes');
        }
        else{
            winMultiContentsView = window.open('/admin/eventmanage/event/v5/template/slide/pop_pcweb_slide.asp?eC='+eventcode+'&menuidx='+menuidx,'imgtxt','width=1450,height=800,scrollbars=yes,resizable=yes');
        }
    }
    else if(menudiv=="9"){
        if(device=="M"){
            winMultiContentsView = window.open('/admin/eventmanage/event/v5/template/addbanner/pop_mobile_addbanner.asp?eC='+eventcode+'&menuidx='+menuidx,'imgtxt','width=1450,height=800,scrollbars=yes,resizable=yes');
        }
        else{
            winMultiContentsView = window.open('/admin/eventmanage/event/v5/template/addbanner/pop_pc_addbanner.asp?eC='+eventcode+'&menuidx='+menuidx,'imgtxt','width=1450,height=800,scrollbars=yes,resizable=yes');
        }
    }
    else if(menudiv=="10"){
        winMultiContentsView = window.open('/admin/eventmanage/event/v5/popup/pop_event_execfile.asp?wm='+device+'&eC='+eventcode+'&menuidx='+menuidx,'customboxinfo','width=1024,height=500,scrollbars=yes,resizable=yes');
    }
    else if(menudiv=="11"){
        winMultiContentsView = window.open('/admin/eventmanage/event/v5/popup/pop_event_imagelink.asp?wm='+device+'&eC='+eventcode+'&menuidx='+menuidx,'customboxinfo','width=1024,height=500,scrollbars=yes,resizable=yes');
    }
    else if(menudiv=="12") {
        winMultiContentsView = window.open('/admin/eventmanage/event/v5/popup/pop_event_customgroupitem2.asp?eC='+eventcode+'&menuidx='+menuidx+'&device='+device,'customgrouptemplateinfo','width=1024,height=800,scrollbars=yes,resizable=yes');
    }
    else if(menudiv=="13") {
        winMultiContentsView = window.open('/admin/eventmanage/event/v5/popup/pop_event_tabbar_template.asp?eC='+eventcode+'&menuidx='+menuidx+'&device='+device,'customgrouptemplateinfo','width=' + window.innerWidth + ',height=800,scrollbars=yes,resizable=yes');
    }
    winMultiContentsView.focus();
}

function fnMultiContentsSortSet(){
    document.frmEvt.target="ifrmProc";
    document.frmEvt.action="/admin/eventmanage/event/v5/popup/multicontentssort_process.asp";
    document.frmEvt.submit();
}

function fnRegItems(eCode, gCode, eChannel){
    if(eCode==""){
        alert("�⺻���� ������ �����ؾ� ���� �����մϴ�.");
    }
    else{
        var wItemsView;
        wItemsView = window.open('/admin/eventmanage/event/v5/popup/eventitem_regist.asp?eC='+eCode+'&selG='+gCode+'&eCh='+eChannel,'itemsreg','width=1400,height=800,scrollbars=yes,resizable=yes');
        wItemsView.focus();
    }
}

function fnGroupManager(eCode, gCode, smode, eChannel){ 
    if(eCode==""){
        alert("�⺻���� ������ �����ؾ� ���� �����մϴ�.");
    }
    else{
        var winG 
        var vYear = "<%=vYear%>";  
        winG = window.open('/admin/eventmanage/event/v5/popup/pop_eventitem_group.asp?yr='+vYear+'&eC='+eCode+'&eGC='+gCode+'&eCh='+eChannel,'popG','width=800, height=800,scrollbars=yes,resizable=yes');
        winG.focus();
    }
}

function fnEventFunction(eCode){ 
    if(eCode==""){
        alert("�⺻���� ������ �����ؾ� ���� �����մϴ�.");
    }
    else{
        var winfunction 
        var vYear = "<%=vYear%>";  
        winfunction = window.open('/admin/eventmanage/event/v5/popup/pop_event_functioninfo.asp?eC='+eCode,'popfunction','width=1024, height=800,scrollbars=yes,resizable=yes');
        winfunction.focus();
    }
}

function fnCultureStationContentsSet(eventcode){
    if(eventcode==""){
        alert("�⺻���� ������ �����ؾ� ���� �����մϴ�.");
    }
    else{
        var winMenuView;
        winMenuView = window.open('/admin/eventmanage/event/v5/popup/pop_culturestation_contentsinfo.asp?eC='+eventcode,'menuinfo','width=1024,height=850,scrollbars=yes,resizable=yes');
        winMenuView.focus();
    }
}

function fnRelationEvent(eventcode){
    if(eventcode==""){
        alert("�⺻���� ������ �����ؾ� ���� �����մϴ�.");
    }
    else{
        var winRelationView;
        winRelationView = window.open('/admin/eventmanage/event/v5/popup/pop_relationEventinfo.asp?eC='+eventcode,'relationinfo','width=1024,height=700,scrollbars=yes,resizable=yes');
        winRelationView.focus();
    }
}

function fnCopyEventSet(){
    var winSelectEvnt;
    winSelectEvnt = window.open('/admin/eventmanage/event/v5/popup/pop_event_select.asp?mode=copy','eventselect','width=900,height=600,scrollbars=yes,resizable=yes');
    winSelectEvnt.focus();
}

function fnPreViewChange(target){
    document.body.scrollIntoView(true);
    <% if ekind="5" then '���Ľ����̼�%>
    if(target=="M"){
        ifrmView.location.href="<%=LinkURL%>/_culturestation/culturestation_event.asp?evt_code=<% =eCode %>";
        document.getElementById("ifrmView").height=600;
        document.getElementById("ifrmView").width=400;
        $("#viewset").val("M");
        $("#momenu").addClass("selected");
        $("#pcmenu").removeClass("selected");
    }
    else{
        ifrmView.location.href="<%=LinkURL2%>/_culturestation/culturestation_event.asp?evt_code=<% =eCode %>";
        document.getElementById("ifrmView").height=800;
        document.getElementById("ifrmView").width=1500;
        $("#viewset").val("P");
        $("#pcmenu").addClass("selected");
        $("#momenu").removeClass("selected");
    }
    <% else %>
    if(target=="M"){
        ifrmView.location.href="<%=LinkURL%>/event/adminView/eventmain.asp?eventid=<% =eCode %>&stepdiv=<% =stepdiv %>";
        document.getElementById("ifrmView").height=600;
        document.getElementById("ifrmView").width=400;
        $("#viewset").val("M");
        $("#momenu").addClass("selected");
        $("#pcmenu").removeClass("selected");
    }
    else{
        ifrmView.location.href="<%=LinkURL2%>/event/adminView/eventmain.asp?eventid=<% =eCode %>&stepdiv=<% =stepdiv %>";
        document.getElementById("ifrmView").height=800;
        document.getElementById("ifrmView").width=1500;
        $("#viewset").val("P");
        $("#pcmenu").addClass("selected");
        $("#momenu").removeClass("selected");
    }
    <% end if %>
}
            
function fnEventStateSet(ecode){
    var estate=document.frmEvt.eventstate.value;
    $.ajax({
        type: "POST",
        url: "/admin/eventmanage/event/v5/lib/ajaxEventStateSet.asp",
        data: "eC="+ecode+"&eState="+estate,
        cache: false,
        success: function(message) {
            if(message=="0") {
                alert("����Ǿ����ϴ�.");
            }
            else if(message=="1"){
                alert("��ȿ���� ���� ������ �Դϴ�. �ٽ� �õ��� �ּ���.");
            }
            else if(message=="2"){
                alert("������ ó���� ������ �߻��Ͽ����ϴ�.");
            }
        },
        error: function(err) {
            alert(err.responseText);
        }
    });
}

function fnMarketingEventPopup(marketingKind){
    let winpickupView;

    switch(marketingKind){
        case 1 : winpickupView = window.open('/admin/eventmanage/joobjoob/index.asp?evt_code=<%=eCode%>&open_date=<%=FormatDate(eSdate,"0000-00-00 00:00:00")%>&end_date=<%=FormatDate(eEdate,"0000-00-00 00:00:00")%>','pickupinfo','width=1024,height=830,scrollbars=yes,resizable=yes');
            break;
        case 2 : winpickupView = window.open('/admin/eventmanage/event/v5/popup/pop_attendance_event_setting.asp?evt_code=<%=eCode%>','attendanceinfo','width=1024,height=900,scrollbars=yes,resizable=yes');
            break;
        case 3 : winpickupView = window.open('/admin/eventmanage/event/v5/popup/pop_login_mileage.asp?evt_code=<%=eCode%>','loginmileageinfo','width=1024,height=400,scrollbars=yes,resizable=yes');
            break;
        case 4 : winpickupView = window.open('/admin/eventmanage/event/v5/popup/pop_event_item_info_link.asp?evt_code=<%=eCode%>','iteminfolink','width=1024,height=400,scrollbars=yes,resizable=yes');
            break;
        case 5 : winpickupView = window.open('/admin/eventmanage/event/v5/popup/pop_app_event_setting.asp?evt_code=<%=eCode%>','appeventlink','width=1024,height=830,scrollbars=yes,resizable=yes');
            break;
        case 6 : winpickupView = window.open('/admin/eventmanage/event/v5/popup/pop_secret_shop_setting.asp?evt_code=<%=eCode%>','secretshoplink','width=1024,height=300,scrollbars=yes,resizable=yes');
            break;
    }

    winpickupView.focus();
}

$(function(){
    //$("#accordion").accordion();
	//�巡��
	$("#MsubList").sortable({
        placeholder: "ui-state-highlight",
		start: function(event, ui) {
			ui.placeholder.html('<li>&nbsp;</li>');
		},
		stop: function(){
			var i=99999;
			$(this).find("input[name^='sort']").each(function(){
				if(i>$(this).val()) i=$(this).val()
			});
			if(i<=0) i=1;
			$(this).find("input[name^='sort']").each(function(){
				$(this).val(i);
				i++;
			});
		}
	});
});
</script>
<script>
$(document).ready(function(){
    $('.evtSidebarV19 .[class^=depth2]').slideUp();
    $(".depth1 > span.mdi").removeClass('mdi-minus').addClass('mdi-plus');
    $(".evtSidebarV19 .depth1").on("click", function(i){
        i.preventDefault();
        if($(this).children('span.mdi').hasClass("mdi-minus")){
            $(this).siblings('.depth2').slideUp();
            $(this).children('span.mdi').removeClass('mdi-minus').addClass('mdi-plus');
        }
        else{
            $(this).siblings('.depth2').slideDown();
            $(this).children('span.mdi').removeClass('mdi-plus').addClass('mdi-minus');
        }
    });
<% if togglediv="1" then %>
    $("#basicmenu .[class^=depth2]").slideToggle();
    $("#basicmenu .depth1 > span.mdi").removeClass('mdi-plus').addClass('mdi-minus');
<% elseif togglediv="2" then %>
    $("#contentsmenu .[class^=depth2]").slideToggle();
    $("#contentsmenu .depth1 > span.mdi").removeClass('mdi-plus').addClass('mdi-minus');
<% elseif togglediv="3" then %>
    $("#itemmenu .[class^=depth2]").slideToggle();
    $("#itemmenu .depth1 > span.mdi").removeClass('mdi-plus').addClass('mdi-minus');
<% elseif togglediv="4" then %>
    $("#functionmenu .[class^=depth2]").slideToggle();
    $("#functionmenu .depth1 > span.mdi").removeClass('mdi-plus').addClass('mdi-minus');
<% elseif togglediv="5" then %>
    $("#multicontentsmenu .[class^=depth2]").slideToggle();
    $("#multicontentsmenu .depth1 > span.mdi").removeClass('mdi-plus').addClass('mdi-minus');
<% end if %>
<% if multiconOpen="Y" then %>
    fnMultiContentsDeviceSet(<% =eCode %>,<% =menuidx %>,<% =menudiv %>,'M');
<% end if %>
});

window.document.domain = "10x10.co.kr";
function fnSlidePreView(moveid){
    window.document.domain = "10x10.co.kr";
    //$("#ifrmView").get(0).contentWindow.fnMoveDivision(moveid);
}

function fnSlidePreViewOut(moveid){
    window.document.domain = "10x10.co.kr";
    //$("#ifrmView").get(0).contentWindow.fnBorderDivisionRemove(moveid);
}

function jsContentsMenuSet(frm){
    frm.target="ifrmProc";
    frm.action="/admin/eventmanage/event/v5/popup/contentsmenu_process.asp";
	frm.submit();
}

function jsDeleteContents(menuidx){
    if(menuidx != ""){
        document.frm.menuidx.value=menuidx;
        document.frm.target="ifrmProc";
        document.frm.submit();
    }
}

function fnCopyEventImageSet(){
    var winSelectEvnt;
    winSelectEvnt = window.open('/admin/eventmanage/event/v5/popup/pop_event_select.asp?mode=imgcopy&eC=<% =eCode %>','eventselect','width=900,height=600,scrollbars=yes,resizable=yes');
    winSelectEvnt.focus();
}

function fnCopyEventImageSet2(ecode){
    $.ajax({
        type: "POST",
        url: "/admin/eventmanage/event/v5/lib/ajaxEventImageCopy.asp",
        data: "eC="+ecode,
        cache: false,
        success: function(message) {
            if(message=="0") {
                alert("�̹��� ���簡 �Ϸ�Ǿ����ϴ�.");
                location.reload();
            }
            else if(message=="1"){
                alert("��ȿ���� ���� ������ �Դϴ�. �ٽ� �õ��� �ּ���.");
            }
            else if(message=="2"){
                alert("������ ó���� ������ �߻��Ͽ����ϴ�.");
            }
        },
        error: function(err) {
            alert(err.responseText);
        }
    });
}

function fnEventReset(ecode){
    $.ajax({
        type: "POST",
        url: "/admin/eventmanage/event/v5/lib/ajaxEventReset.asp",
        data: "eC="+ecode,
        cache: false,
        success: function(message) {
            if(message=="0") {
                alert("�ʱ�ȭ �Ϸ�Ǿ����ϴ�.");
                location.reload();
            }
            else if(message=="1"){
                alert("��ȿ���� ���� ������ �Դϴ�. �ٽ� �õ��� �ּ���.");
            }
            else if(message=="2"){
                alert("������ ó���� ������ �߻��Ͽ����ϴ�.");
            }
        },
        error: function(err) {
            alert(err.responseText);
        }
    });
}
</script>

<form name="frmEvt" method="post" style="margin:0px;">
<input type="hidden" name="evt_code" value="<%=eCode%>"/>
<input type="hidden" name="viewset" id="viewset" value="<%=viewset%>"/>
<input type="hidden" name="imod" value="MI">
<div class="contentWrapV19">
    <!-- sidebar -->
    <ul class="evtSidebarV19">
        <li class="btnPreivew">
            <div class="depth1"><i class="mdi mdi-eye"></i> �̸�����</div>
            <div class="tPad10 bPad10 bgWht1">
                <button class="btn4 btnGrey1 lMar05">PC</button>
                <button class="btn4 btnGrey1 lMar05">MW</button>
            </div>
        </li>
        <li id="basicmenu">
            <div class="depth1"><i class="mdi mdi-settings"></i> �⺻���� <span class="mdi mdi-minus"></span></div>
            <ul class="depth2">
            <% if fnCheckAuthImageCopy(session("ssBctId")) and estate < 4 then %>
                <li class="tPad10 bPad10 bgGry2 ct"><button class="btn4 btnWhite1" onclick="fnCopyEventImageSet('<%=eCode%>');return false;">�̹��� �����ϱ�</button></li>
                <div class="depth3Wrap">
                    <table class="depth3">
                        <% if EvtCopyCode > 0 then %>
                        <tr>
                            <td>�����̺�Ʈ �ڵ� : <% = EvtCopyCode %></td>
                        </tr>
                        <% end if %>
                        <% if ImgCopyCode > 0 then %>
                        <tr>
                            <td>�̹������� ����̺�Ʈ �ڵ� : <% = ImgCopyCode %></td>
                        </tr>
                        <% end if %>
                    </table>
                </div>
            <% end if %>
            <% if (C_ADMIN_AUTH or C_MD_AUTH) and ((estate=7 or estate=9) and eEdate <= Now()) then %>
                <li class="tPad10 bPad10 bgGry2 ct"><button class="btn4 btnWhite1" onclick="fnEventReset('<%=eCode%>');return false;">�̺�Ʈ �ʱ�ȭ</button></li>
            <% end if %>
            <% if eCode<>"" then %>
                <li>
                    <p><span>��ȹ�� �ڵ�</span><strong class="evtCode"><%=eCode%></strong></p>
                </li>
                <li>
                    <!-- for dev msg �ʼ� �Է� ��ҿ� essential Ŭ���� �߰� -->
                    <p onclick="fnBasicInfoSet('<%=eCode%>');"><span class="essential">�⺻���� ����</span><span class="mdi mdi-chevron-right cGy3"></span></p>
                    <div class="depth3Wrap">
                        <table class="depth3" onclick="fnBasicInfoSet('<%=eCode%>');">
                            <tr>
                                <th>����</th>
                                <td><% GetEvnetKindName "eventkind", ekind %></td>
                            </tr>
                            <% if ekind="5" then %>
                            <tr>
                                <th>����</th>
                                <td><% if evt_type="0" then %>������<% else %>�о��<% end if %></td>
                            </tr>
                            <% end if %>
                            <tr>
                                <th>ä��</th>
                                <td><% if isWeb then %> PC &#10072;<% end if %><% if isMobile then %> Mobile &#10072;<% end if %><% if isApp then %> APP <% end if %></td>
                            </tr>
                            <tr onMouseOver="fnSlidePreView('title')">
                                <th><% if ekind<>"5" then %>����ī�� (����)<% else %>����<% end if %></th>
                                <td><%=ename%>&nbsp;</td>
                            </tr>
                            <% if ekind="5" then %>
                            <tr>
                                <th>��÷����</th>
                                <td><% =subcopyK %>&nbsp;</td>
                            </tr>
                            <% end if %>
                            <tr onMouseOver="fnSlidePreView('edate')">
                                <th>�Ⱓ</th>
                                <td>
                                    <% if eSdate<>"" then %><%=FormatDate(eSdate,"0000.00.00")%> ~ <%=FormatDate(eEdate,"0000.00.00")%><% end if %>
                                    <% if eDateView then %>
                                    <p class="cGy1 fs12">*�Ⱓ ���� ����</p>
                                    <% end if %>
                                </td>
                            </tr>
                            <% if ekind<>"5" then %>
                            <tr>
                                <th>Ÿ��</th>
                                <td><% if esale then %> ���� &#10072;<% end if %><% if egift then %> ����ǰ &#10072;<% end if %><% if ecoupon then %> ���� &#10072;<% end if %><% if eonlyten then %> Only-TenByTen &#10072;<% end if %><% if eOneplusone then %> 1+1 &#10072;<% end if %><% if eFreedelivery then %> ������ &#10072;<% end if %><% if eBookingsell then %> �����Ǹ� &#10072;<% end if %><% if eDiary then %> DiaryStory &#10072;<% end if %><% if eNew then %> ��Ī &#10072;<% end if %></td>
                            </tr>
                            <% end if %>
                            <% if ecomment or ebbs or eitemps or eisblogurl then %>
                            <tr>
                                <th>���</th>
                                <td>
                                    <% if ecomment then %><span>�ڸ�Ʈ</span><% end if %>
                                    <% if ebbs then %><span>�����ڸ�Ʈ</span><% end if %>
                                    <% if eitemps then %><span>��ǰ�ı�</span><% end if %>
                                </td>
                            </tr>
                            <% end if %>
                            <% if ePdate<>"" then %>
                            <tr>
                                <th>��÷ ��ǥ</th>
                                <td><%=FormatDate(ePdate,"0000.00.00")%></td>
                            </tr>
                            <% end if %>
                            <% if ekind<>"5" then %>
                            <tr>
                                <th>�߿䵵</th>
                                <td><% if elevel="1" then %><font color="red">�ֻ�</font><% elseif elevel="2" then %>��<% elseif elevel="3" then %>��<% else %>��<% end if %></td>
                            </tr>
                            <tr>
                                <th>��������</th>
                                <td><%=estimateSalePrice%></td>
                            </tr>
                            <tr>
                                <th>��ü</th>
                                <td><% if eman="1" then %>10x10<% elseif eman="2" then %>��ü<% end if %></td>
                            </tr>
                            <% end if %>
                            <tr>
                                <th>�������</th>
                                <td><% if eusing<>"" then %><% if eusing="Y" then %>YES<% else %>NO<% end if %><% end if %></td>
                            </tr>
                            <% if esale then %>
                            <tr>
                                <th>����</th>
                                <td class="cRd2"><% if salePer<>"" then %><strong>~<%=salePer%>%</strong><% else %>�׷������ ���� �� �ڵ� üũ<% end if %></td>
                            </tr>
                            <% end if %>
                            <% if ecoupon then %>
                            <tr>
                                <th>����</th>
                                <td class="cGn2"><% if saleCPer<>"" then %><strong>~<%=saleCPer%>%</strong><% else %>�׷������ ���� �� �ڵ� üũ<% end if %></td>
                            </tr>
                            <% end if %>
                        </table>
                    </div>
                </li>
                <li>
                    <p onclick="fnWorkerInfoSet('<%=eCode%>');"><span>����� ����</span><span class="mdi mdi-chevron-right cGy3"></span></p>
                    <div class="depth3Wrap">
                        <table class="depth3" onclick="fnWorkerInfoSet('<%=eCode%>');">
                            <tr>
                                <th>��ȹ</th>
                                <td><%=eMDnm%></td>
                            </tr>
                            <tr>
                                <th>�����̳�</th>
                                <td><%=edgnm%>&nbsp;</td>
                            </tr>
                            <tr>
                                <th>�ۺ���</th>
                                <td><%=epsnm%>&nbsp;</td>
                            </tr>
                            <tr>
                                <th>����</th>
                                <td><%=edpnm%>&nbsp;</td>
                            </tr>
                            <% if sWorkTag<>"" then %>
                            <tr>
                                <th>�۾�����</th>
                                <td><%=sWorkTag%>&nbsp;</td>
                            </tr>
                            <% end if %>
                        </table>
                        <table class="depth3">
                            <tr class="topLineGrey2" onclick="fnViewWorkMemo();">
                                <th class="tPad10">�۾����޻��� <span class="mdi mdi-open-in-new cGy3 fs16"></span></th>
                                <td class="tPad10"><div class="workerMsg pad10 bgWht1"><%=efwd%>&nbsp;</div></td>
                            </tr>
                        </table>
                    </div>
                </li>
                <% if ekind<>"5" then %>
                <li onclick="fnListBannerSet('<%=eCode%>');">
                    <p><span class="essential">�⺻��� ����</span><span class="mdi mdi-chevron-right cGy3"></span></p>
                    <div class="depth3Wrap bPad15">
                        <% if eEtcitemimg<>"" then %>
                        <span class="previewThumb80H"><img src="<%=eEtcitemimg%>" alt=""></span>
                        <% end if %>
                        <% if ebimgMo2014<>"" then %>
                        <span class="previewThumb80H lMar05"><img src="<%=ebimgMo2014%>" alt=""></span>
                        <% end if %>
                    </div>
                </li>
                <% end if %>
                <% if ekind<>"5" then %>
                <li onclick="fnEventInfoSet('<%=eCode%>');">
                    <p><span class="essential">��ȹ�� ����</span><span class="mdi mdi-chevron-right cGy3"></span></p>
                    <div class="depth3Wrap">
                        <table class="depth3">
                            <% if nocate<>"Y" then %>
                            <tr>
                                <th>ī�װ�</th>
                                <td><%=DispCateName%>&nbsp;</td>
                            </tr>
                            <% end if %>
                            <tr>
                                <th>�귣��</th>
                                <td><%=ebrand%>&nbsp;</td>
                            </tr>
                            <tr>
                                <th>TAG</th>
                                <td><%=etag%>&nbsp;</td>
                            </tr>
                        </table>
                    </div>
                </li>
                <% end if %>
                <li onclick="fnEventSnsShareSet('<%=eCode%>')">
                    <p><span class="essential">SNS ���� ����</span></p>
                    <div class="depth3Wrap">
                        <table class="depth3">
                            <tr>
                                <th>īī����</th>
                                <td><%=kakaoTitle%></td>
                            </tr>
                        </table>
                    </div>
                </li>
                <% if ekind="28" and marketing_event_kind=1 then %>
                    <li onclick="fnMarketingEventPopup(1);">
                        <p><span class="essential">���� ����</span><span class="mdi mdi-chevron-right cGy3"></span></p>
                    </li>
                <% end if %>
                <% if ekind="28" and marketing_event_kind=2 then %>
                    <li onclick="fnMarketingEventPopup(2);">
                        <p><span class="essential">�⼮üũ ����</span><span class="mdi mdi-chevron-right cGy3"></span></p>
                    </li>
                <% end if %>
                <% if ekind="28" and marketing_event_kind=3 then %>
                    <li onclick="fnMarketingEventPopup(3);">
                        <p><span class="essential">�α��� ���ϸ��� ����</span><span class="mdi mdi-chevron-right cGy3"></span></p>
                    </li>
                <% end if %>
                <% if ekind="28" and marketing_event_kind=4 then %>
                    <li onclick="fnMarketingEventPopup(4);">
                        <p><span class="essential">��ǰ ���� ����</span><span class="mdi mdi-chevron-right cGy3"></span></p>
                    </li>
                <% end if %>
                <% if ekind="28" and marketing_event_kind=5 then %>
                    <li onclick="fnMarketingEventPopup(5);">
                        <p><span class="essential">������ ������ ����</span><span class="mdi mdi-chevron-right cGy3"></span></p>
                    </li>
                <% end if %>
                <% if ekind="28" and marketing_event_kind=6 then %>
                    <li onclick="fnMarketingEventPopup(6);">
                        <p><span class="essential">����� �� ����</span><span class="mdi mdi-chevron-right cGy3"></span></p>
                    </li>
                <% end if %>
            <% else %>
                <li onclick="fnCopyEventSet('<%=eCode%>');">
                    <p><span>��ȹ�� �����ϱ�</span><span class="mdi mdi-chevron-right"></span></p>
                </li>
                <li onclick="fnBasicInfoSet('<%=eCode%>');">
                    <p><span class="essential">�⺻���� ����</span><span class="mdi mdi-chevron-right"></span></p>
                <li>
            <% end if %>
                <!--// �̺�Ʈ ������ ���� -->
            </ul>
        </li>
        <li id="contentsmenu">
            <div class="depth1"><i class="mdi mdi-image fs20" style="vertical-align:-2px;"></i> ������ ���� <span class="mdi mdi-minus"></span></div>
            <ul class="depth2">
                <% if ekind<>"13" and ekind<>"5" then %>
                <li onclick="fnEventTopInfoSet('<%=eCode%>');">
                    <p><span class="essential">���Ÿ��Ʋ</span><span class="mdi mdi-chevron-right"></span></p>
                    <div class="depth3Wrap">
                        <% if etemp_mo<>"10" then %>
                        <strong>Mobile</strong>
                        <table class="depth3 tMar10">
                            <tr>
                                <th>����ī��</th>
                                <td><%=title_mo%>&nbsp;</td>
                            </tr>
                            <tr>
                                <th>����ī��</th>
                                <td><%=enamesub%>&nbsp;</td>
                            </tr>
                            <tr>
                                <th>�Է¹��</th>
                                <td><% if etemp_mo="11" then %>���ø����<% elseif etemp_mo="11" then %>Multi3��<% else %>���۾�<% end if %>&nbsp;</td>
                            </tr>
                            <tr>
                                <th>����</th>
                                <td><% if mdbntypemo="D" then %>�̹��� �����̵�<% else %>����÷�<% end if %>&nbsp;</td>
                            </tr>
                            <tr>
                                <th>����÷�</th>
                                <td>
                                    <div class="colorPicker">
                                        <ul>
                                            <li><span style="<%=fnEventBarColorCode(themecolormo)%>"></span></li>
                                        </ul>
                                    </div>
                                </td>
                            </tr>
                        </table>

                        <strong>PC</strong>
                        <table class="depth3 tMar10">
                            <tr>
                                <th>����ī��</th>
                                <td><%=title_pc%>&nbsp;</td>
                            </tr>
                            <tr>
                                <th>����ī��</th>
                                <td><%=subcopyK%>&nbsp;</td>
                            </tr>
                            <tr>
                                <th>�Է¹��</th>
                                <td><% if etemp="10" then %>���ø����<% else %>���۾�<% end if %>&nbsp;</td>
                            </tr>
                            <tr>
                                <th>����</th>
                                <td><% if mdbntype="D" then %>�̹��� �����̵�<% else %>����÷�<% end if %></td>
                            </tr>
                            <tr>
                                <th>����÷�</th>
                                <td>
                                    <div class="colorPicker">
                                        <ul>
                                            <li><span style="<%=fnEventBarColorCode(themecolor)%>"></span></li>
                                        </ul>
                                    </div>
                                </td>
                            </tr>
                        </table>
                        <% end if %>
                    </div>
                </li>
                <% end if %>
                <% if ekind="5" then %>
                <li onclick="fnCultureStationContentsSet('<%=eCode%>');">
                    <p><span class="essential">���Ÿ��Ʋ</span><span class="mdi mdi-chevron-right"></span></p>
                    <div class="depth3Wrap">
                        <table class="depth3 tMar10">
                            <tr>
                                <th>�����ü</th>
                                <td><% =CommentTitle %>&nbsp;</td>
                            </tr>
                            <tr>
                                <th>������</th>
                                <td>
                                    <span class="previewThumb80W"><img src="<% =evt_mainIMG %>" alt=""></span>
                                </td>
                            </tr>
                            <tr>
                                <th>����÷�</th>
                                <td>
                                    <div class="colorPicker">
                                        <ul>
                                            <li><span style="<%=fnEventBarColorCode(themecolor)%>"></span></li>
                                        </ul>
                                    </div>
                                </td>
                            </tr>
                        </table>
                    </div>
                </li>
                <% end if %>
                <% if ekind="5" or etemp_mo="10" then %>
                <% else %>
                <li onclick="fnGiftInfoSet('<%=eCode%>');return false;">
                    <p><span>GIFT</span><span class="mdi mdi-chevron-right"></span></p>
                </li>
                <% end if %>
            </ul>
        </li>
        <% if ekind="13" then %>
        <% else %>
        <li id="multicontentsmenu">
            <div class="depth1"><i class="mdi mdi-image fs20" style="vertical-align:-2px;"></i> ���ø� �߰� ������ ���� <span class="mdi mdi-minus"></span></div>
            <ul class="depth2">
                <li class="tPad10 bPad10 lPad15 bgGry2">
                    <select name="menudiv" class=" brderGry1" style="width:70%;">
                        <option value="" disabled selected>�߰� ������</option>
                        <option value="1">�̹��� �����̵�</option>
                        <option value="2">����</option>
                        <% if ekind<>"5" then %>
                        <option value="3">�귣�彺�丮</option>
                        <option value="4">��õ ����Ʈ</option>
                        <% end if %>
                        <option value="5">������ ����</option>
                        <% if ekind<>"5" then %>
                        <option value="6">��� �����̵�</option>
                        <option value="7">�̹��� & HTML</option>
                        <option value="8">�̹��� ���ø� �����̵�</option>
                        <option value="9">������</option>
                        <option value="11">�̹�����ũ</option>
                        <option value="12">��ǰ ���� ����</option>
                        <option value="13">�ǹ�</option>
                        <% end if %>
                    </select>
                    <button class="btn4 btnBlue1 lMar05" onclick="jsContentsMenuSet(this.form);return false;">�߰�</button>
                </li>
            </ul>
            <ul class="depth2" id="MsubList">
                <% if ekind<>"13" then %>
                <% If isArray(ArrcMultiContentsMenu) Then %>
                <li class="tPad10 bPad10 bgGry2 ct">����� / APP</li>
                <% For ix = 0 To UBound(ArrcMultiContentsMenu,2) %>
                <li>
                    <% If ArrcMultiContentsMenu(1,ix)="12" Then '// ��ǰ ���� ���� %>
                        <p style="padding:5px 15px;">
                            <span class="mdi mdi-equal rMar10"></span>
                            <span onclick="fnMultiContentsDeviceSet('<%=eCode%>','<%=ArrcMultiContentsMenu(0,ix)%>','<%=ArrcMultiContentsMenu(1,ix)%>','M');return false;" onMouseOver="fnSlidePreView('multi<%=ArrcMultiContentsMenu(0,ix)%>')" onMouseOut="fnSlidePreViewOut('multi<%=ArrcMultiContentsMenu(0,ix)%>')" style="vertical-align:middle;color:<% if ArrcMultiContentsMenu(5,ix) > 0 then %>#4075ff<% else %><% end if %>">
                                <%=GetMenuDivName(ArrcMultiContentsMenu(1,ix))%>
                            </span>
                            <input type="hidden" name="sort" value="<%=ArrcMultiContentsMenu(2,ix)%>"/>
                            <input type="hidden" name="idx" value="<%=ArrcMultiContentsMenu(0,ix)%>"/>
                            <button class="btn4 btnGrey1 lMar10" onclick="jsDeleteContents(<%=ArrcMultiContentsMenu(0,ix)%>);return false;">����</button>
                            <% If ArrcMultiContentsMenu(6,ix) <> "" Then %>
                                <img src="<%=ArrcMultiContentsMenu(6,ix)%>" style="height:26px;margin-left:10px;">
                            <% End If %>
                            <span class="mdi mdi-chevron-right"></span>
                        </p>
                    <% Else %>
                        <p>
                            <span class="mdi mdi-equal rMar10"></span>
                            <% if ArrcMultiContentsMenu(1,ix)="1" or ArrcMultiContentsMenu(1,ix)="6" or ArrcMultiContentsMenu(1,ix)="7" or ArrcMultiContentsMenu(1,ix)="8" or ArrcMultiContentsMenu(1,ix)="9" or ArrcMultiContentsMenu(1,ix)="10" or ArrcMultiContentsMenu(1,ix)="11" or ArrcMultiContentsMenu(1,ix)="13" then %>
                            <span onclick="fnMultiContentsDeviceSet('<%=eCode%>','<%=ArrcMultiContentsMenu(0,ix)%>','<%=ArrcMultiContentsMenu(1,ix)%>','M');return false;" onMouseOver="fnSlidePreView('multi<%=ArrcMultiContentsMenu(0,ix)%>')" onMouseOut="fnSlidePreViewOut('multi<%=ArrcMultiContentsMenu(0,ix)%>')" style="color:<% if ArrcMultiContentsMenu(4,ix) > 0 then %>#4075ff<% else %><% end if %>">
                            <% else %>
                            <span onclick="fnMultiContentsSet('<%=eCode%>','<%=ArrcMultiContentsMenu(0,ix)%>','<%=ArrcMultiContentsMenu(1,ix)%>');return false;" onMouseOver="fnSlidePreView('multi<%=ArrcMultiContentsMenu(0,ix)%>')" onMouseOut="fnSlidePreViewOut('multi<%=ArrcMultiContentsMenu(0,ix)%>')" style="color:<% if ArrcMultiContentsMenu(4,ix) > 0 or ArrcMultiContentsMenu(5,ix) > 0 then %>#4075ff<% else %><% end if %>">
                            <% end if %>
                            <%=GetMenuDivName(ArrcMultiContentsMenu(1,ix))%></span>
                            <input type="hidden" name="sort" value="<%=ArrcMultiContentsMenu(2,ix)%>"/><input type="hidden" name="idx" value="<%=ArrcMultiContentsMenu(0,ix)%>"/><button class="btn4 btnGrey1 lMar10" onclick="jsDeleteContents(<%=ArrcMultiContentsMenu(0,ix)%>);return false;">����</button>
                            <span class="mdi mdi-chevron-right"></span>
                        </p>
                    <% End If %>
                </li>
                <% Next %>
                <% End If %>
                <% End If %>
            </ul>
            <ul class="depth2">
                <% if ekind<>"13" then %>
                <% If isArray(ArrcMultiContentsMenu) Then %>
                <li class="tPad10 bPad10 bgGry2 ct">PCWeb</li>
                <% For ix = 0 To UBound(ArrcMultiContentsMenu,2) %>
                <li>
                    <% If ArrcMultiContentsMenu(1,ix)="12" Then '// ��ǰ ���� ���� %>
                        <p style="padding:5px 15px;">
                            <span class="mdi mdi-equal rMar10"></span>
                            <span onclick="fnMultiContentsDeviceSet('<%=eCode%>','<%=ArrcMultiContentsMenu(0,ix)%>','<%=ArrcMultiContentsMenu(1,ix)%>','W');return false;" onMouseOver="fnSlidePreView('multi<%=ArrcMultiContentsMenu(0,ix)%>')" onMouseOut="fnSlidePreViewOut('multi<%=ArrcMultiContentsMenu(0,ix)%>')" style="color:<% if ArrcMultiContentsMenu(5,ix) > 0 then %>#4075ff<% else %><% end if %>">
                            <%=GetMenuDivName(ArrcMultiContentsMenu(1,ix))%></span>
                            <button class="btn4 btnGrey1 lMar10" onclick="jsDeleteContents(<%=ArrcMultiContentsMenu(0,ix)%>);return false;">����</button>
                            <% If ArrcMultiContentsMenu(7,ix) <> "" Then %>
                                <img src="<%=ArrcMultiContentsMenu(7,ix)%>" style="height:26px;margin-left:10px;">
                            <% End If %>
                            <span class="mdi mdi-chevron-right"></span>
                        </p>
                    <% Else %>
                        <p><span class="mdi mdi-equal rMar10"></span>
                        <% if ArrcMultiContentsMenu(1,ix)="1" or ArrcMultiContentsMenu(1,ix)="6" or ArrcMultiContentsMenu(1,ix)="7" or ArrcMultiContentsMenu(1,ix)="8" or ArrcMultiContentsMenu(1,ix)="9" or ArrcMultiContentsMenu(1,ix)="10" or ArrcMultiContentsMenu(1,ix)="11" or ArrcMultiContentsMenu(1,ix)="12" or ArrcMultiContentsMenu(1,ix)="13" then %>
                        <span onclick="fnMultiContentsDeviceSet('<%=eCode%>','<%=ArrcMultiContentsMenu(0,ix)%>','<%=ArrcMultiContentsMenu(1,ix)%>','W');return false;" onMouseOver="fnSlidePreView('multi<%=ArrcMultiContentsMenu(0,ix)%>')" onMouseOut="fnSlidePreViewOut('multi<%=ArrcMultiContentsMenu(0,ix)%>')" style="color:<% if ArrcMultiContentsMenu(5,ix) > 0 then %>#4075ff<% else %><% end if %>">
                        <% else %>
                        <span onclick="fnMultiContentsSet('<%=eCode%>','<%=ArrcMultiContentsMenu(0,ix)%>','<%=ArrcMultiContentsMenu(1,ix)%>');return false;" onMouseOver="fnSlidePreView('multi<%=ArrcMultiContentsMenu(0,ix)%>')" onMouseOut="fnSlidePreViewOut('multi<%=ArrcMultiContentsMenu(0,ix)%>')" style="color:<% if ArrcMultiContentsMenu(4,ix) > 0 or ArrcMultiContentsMenu(5,ix) > 0 then %>#4075ff<% else %><% end if %>">
                        <% end if %>
                        <%=GetMenuDivName(ArrcMultiContentsMenu(1,ix))%></span>
                        <button class="btn4 btnGrey1 lMar10" onclick="jsDeleteContents(<%=ArrcMultiContentsMenu(0,ix)%>);return false;">����</button><span class="mdi mdi-chevron-right"></span></p>
                    <% End If %>
                </li>
                <% Next %>
                <% End If %>
                <% End If %>
            </ul>
            <ul class="depth2">
                <li class="tPad10 bPad10 bgGry2 ct"><button class="btn4 btnWhite1" onclick="fnMultiContentsSortSet('<%=eCode%>');return false;">������ ���� ����</button></li>
            </ul>
        </li>
        <% End If %>
        <% if ekind="5" or etemp_mo="10" then %>
        <% else %>
        <li id="itemmenu">
            <div class="depth1"><i class="mdi mdi-format-list-bulleted fs20" style="vertical-align:-2px;"></i> ��ǰ ����Ʈ <span class="mdi mdi-minus"></span></div>
            <ul class="depth2">
                <li onclick="fnGroupManager('<%=eCode%>','','I','P');return false;">
                    <p><span>�׷����</span><span class="mdi mdi-chevron-right"></span></p>
                </li>
                <li onclick="fnRegItems('<%=eCode%>','','');return false;">
                    <p><span class="essential">��ǰ����</span><span class="mdi mdi-chevron-right"></span></p>
                </li>
            </ul>
        </li>
        <% End If %>
        <% if ekind="13" or etemp_mo="10" then %>
        <% else %>
        <li id="functionmenu">
            <div class="depth1"><i class="mdi mdi-tooltip-text-outline fs20" style="vertical-align:-2px;"></i> ��� ���� <span class="mdi mdi-minus"></span></div>
            <ul class="depth2">
                <li onclick="fnEventFunction('<%=eCode%>','','');return false;">
                    <p><span class="essential">��ȹ�� ���</span><span class="mdi mdi-chevron-right"></span></p>
                </li>
                <% if ekind="13" or ekind="5" or etemp_mo="10" then %>
                <% else %>
                <li onclick="fnRelationEvent('<%=eCode%>');return false;">
                    <p><span>��õ ��ȹ��</span><span class="mdi mdi-chevron-right"></span></p>
                </li>
                <% End If %>
            </ul>
        </li>
        <% End If %>
    </ul>
    <!--// sidebar -->

    <!-- content -->
    <div class="contentV19">
        <div style="padding:50px; font-size:30px;">
            <div class="tabV19">
                <ul>
                    <li<% if isMobile or isApp then %> class="selected"<% end if %> id="momenu"><a href="javascript:fnPreViewChange('M');">Mobile / App</a></li>
                    <li<% if isWeb then %> class="selected"<% end if %> id="pcmenu"><a href="javascript:fnPreViewChange('P');">PC</a></li>
                </ul>
            </div>
        </div>
        <div style="padding-left:50px; font-size:30px;">
            <% if eCode <> "" then %>
            <% if ekind="5" then '���Ľ����̼�%>
                <iframe name="ifrmView" id="ifrmView" src="<%=LinkURL%>/_culturestation/culturestation_event.asp?evt_code=<%=eCode %>" frameborder="1" width="1500" height="800" style="background-color: #FFFFFF"></iframe>
            <% else %>
                <iframe name="ifrmView" id="ifrmView" src="<%=LinkURL3%>/event/adminView/eventmain.asp?eventid=<%=eCode %>&stepdiv=<%=stepdiv %>" frameborder="1" width="1500" height="800" style="background-color: #FFFFFF"></iframe>
            <% end if %>
            <% end if %>
        </div>
    </div>
    <!--// content -->
</div>
</form>
<form name="frm" method="post" style="margin:0px;" action="/admin/eventmanage/event/v5/popup/contentsmenu_process.asp">
<input type="hidden" name="evt_code" value="<%=eCode%>">
<input type="hidden" name="imod" value="MD">
<input type="hidden" name="menuidx">
</form>
<iframe name="ifrmProc" id="ifrmProc" frameborder="0" width="0" height="0"></iframe>
<script>
$(document).ready(function(){
    $("#disp1").attr("disabled", true);
    $("#disp2").attr("disabled", true);
<% if viewset="M" then %>
    fnPreViewChange('M');
<% else %>
    fnPreViewChange('P');
<% end if %>
});
</script>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
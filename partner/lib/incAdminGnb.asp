<!-- #include virtual="/lib/classes/adminpartner/menuCls.asp" -->
<%
'###########################################################
' Description : ��� gnb
' Hieditor : 2016.11.24 ������ ����
'			 2016.12.27 �ѿ�� ����(menupos �Ķ��Ÿ ��ũ ��� ����)
'###########################################################
%>
<%
'' ����(DIY), �Ϲݾ�ü ������.

Dim menupos, menuposidx '�޴� ��ȣ
Dim clsMenu, conIsOffShopOpen
Dim conParentSize, conParentMenuName,conChildSize,conChildMenuLinkUrl, conChildMenuName, conChildMenuSize, conMenuID, conNotice, conHelp
Dim conPLoop,conLoop
Dim isOffUpBeaExists


Dim ssBctDiv_UsercDiv :ssBctDiv_UsercDiv = session("ssBctDiv")&"_"&session("ssUserCDiv")
Dim ssBctDiv_IsAcademy :ssBctDiv_IsAcademy = (ssBctDiv_UsercDiv="9999_14") or (ssBctDiv_UsercDiv="9999_15")
dim sMenuDiv : sMenuDiv=9999
if ssBctDiv_IsAcademy then	sMenuDiv = 9000

if session("chkOffShop") <> 1 then
    set clsMenu = new CMenuList
    clsMenu.FRectMakerID = session("ssBctID")
    session("isOffshop") = clsMenu.fnChkOffShop(isOffUpBeaExists)
    session("isOffUpBeaExists") = isOffUpBeaExists   ''2016/06/27 �߰�
    set clsMenu = nothing

    session("chkOffShop") = 1       ''2016/06/27 �߰�
end if
conIsOffShopOpen = session("isOffshop")
isOffUpBeaExists = session("isOffUpBeaExists")

'//ž�޴� ����Ʈ ��������=====================================================
Function fnGetTopMenu
if (application("Svr_Info")	= "Dev") then
'response.write "reloadMnu"

end if

	Dim clsTopMenu ,intPLoop,intLoop , intMaxLoop
	Dim clsMenu,  arrParentMenuName(),arrChildMenuLinkUrl(), arrChildMenuName(), arrChildMenuSize(), arrMenuId(),arrNotice(), arrHelp()



	set clsTopMenu = new CMenuList
	 clsTopMenu.FRectUserDiv = sMenuDiv
	 clsTopMenu.sbMenuList

	redim arrParentMenuName(clsTopMenu.FParentSize) 	'--�����޴���
	redim arrChildMenuSize(clsTopMenu.FParentSize) 		'--�����޴��� ���� ���� �޴� ��
	redim arrChildMenuLinkUrl(clsTopMenu.FParentSize,0) '--�����޴� ��ũ
	redim arrChildMenuName(clsTopMenu.FParentSize,0) 	'--�����޴���
	redim arrMenuId(clsTopMenu.FParentSize,0) 			'--�޴����̵�
	redim arrNotice(clsTopMenu.FParentSize,0)
	redim arrHelp(clsTopMenu.FParentSize,0)
	intMaxLoop = 0

	For intPLoop = 0 To clsTopMenu.FParentSize
		arrParentMenuName(intPLoop)  = clsTopMenu.FParentMenuName(intPLoop)
		arrChildMenuSize(intPLoop) = clsTopMenu.FChildSize(intPLoop)

		if intMaxLoop < clsTopMenu.FChildSize(intPLoop) Then '�迭�� �ִ������ ���ϱ�
			intMaxLoop = clsTopMenu.FChildSize(intPLoop)
		end if
		 redim preserve arrChildMenuLinkUrl(clsTopMenu.FParentSize,intMaxLoop)
		 redim preserve arrChildMenuName(clsTopMenu.FParentSize,intMaxLoop)
		 redim preserve arrMenuId(clsTopMenu.FParentSize,intMaxLoop)
		 redim preserve arrNotice(clsTopMenu.FParentSize,intMaxLoop)
		 redim preserve arrHelp(clsTopMenu.FParentSize,intMaxLoop)

    if clsTopMenu.FChildSize(intPLoop) >= 0 then '' = �߰�
		For intLoop = 0 To clsTopMenu.FChildSize(intPLoop)
		 arrChildMenuLinkUrl(intPLoop,intLoop) = clsTopMenu.FChildMenu(intPLoop,intLoop).Flinkurl
		 arrChildMenuName(intPLoop,intLoop) = clsTopMenu.FChildMenu(intPLoop,intLoop).Fmenuname

		 arrMenuId(intPLoop,intLoop) = clsTopMenu.FChildMenu(intPLoop,intLoop).Fid
		 arrNotice(intPLoop,intLoop) = clsTopMenu.FChildMenu(intPLoop,intLoop).Fmenuposnotice
		 arrHelp(intPLoop,intLoop) = clsTopMenu.FChildMenu(intPLoop,intLoop).Fmenuposhelp
		Next
	end if
	Next

set clsTopMenu = nothing

'application ���� ����---------------
	Application.lock
	Application("topParentMenu"&ssBctDiv_UsercDiv) 	= arrParentMenuName
	Application("topChildMenu"&ssBctDiv_UsercDiv) 	= arrChildMenuName
	Application("topMenuID"&ssBctDiv_UsercDiv) 		= arrMenuId
	Application("topChildMenuLink"&ssBctDiv_UsercDiv) = arrChildMenuLinkUrl
	Application("topChildMenuSize"&ssBctDiv_UsercDiv) = arrChildMenuSize
	Application("topNotice"&ssBctDiv_UsercDiv) 		= arrNotice
	Application("topHelp"&ssBctDiv_UsercDiv) 		 	= arrHelp
	Application("chkMenu"&ssBctDiv_UsercDiv) = "1"
	Application.unlock
'-----------------------------------
End Function
'//=============================================================================

'�޴� ��ȭ �������� �Լ� �ҷ�����
'if (Application("chkMenu"&ssBctDiv_UsercDiv) <> "1") then ''(true) or
	call fnGetTopMenu
'end if
	conParentSize = ubound(Application("topParentMenu"&ssBctDiv_UsercDiv))
	conChildSize  = ubound(Application("topChildMenu"&ssBctDiv_UsercDiv),2)
	redim conParentMenuName(conParentSize) 	'--�����޴���
	redim conChildMenuSize(conParentSize)
	redim conChildMenuLinkUrl(conParentSize,conChildSize) '--�����޴� ��ũ
	redim conChildMenuName(conParentSize,conChildSize) 	'--�����޴���
	redim conMenuID(conParentSize,conChildSize)
	redim conNotice(conParentSize,conChildSize)
	redim conHelp(conParentSize,conChildSize)

	conParentMenuName	= Application("topParentMenu"&ssBctDiv_UsercDiv)
	conChildMenuName 	= Application("topChildMenu"&ssBctDiv_UsercDiv)
	conChildMenuLinkUrl= Application("topChildMenuLink"&ssBctDiv_UsercDiv)
	conChildMenuSize	= Application("topChildMenuSize"&ssBctDiv_UsercDiv)
	conMenuID			= Application("topMenuID"&ssBctDiv_UsercDiv)
	conNotice			= Application("topNotice"&ssBctDiv_UsercDiv)
	conHelp			= Application("topHelp"&ssBctDiv_UsercDiv)

dim conCurrent
menupos = requestCheckVar(Request("menupos"),10)  '�޴���ȣ
		
					
%>
<script>
$(function() {
	/*
	var swiper = new Swiper('.gnbWrap .swiper-container', {
		slidesPerView: 'auto',
		spaceBetween:0,
		grabCursor: true,
		scrollbar:'.gnbWrap .swiper-scrollbar'
	});
	*/
});
</script>
<style>
.swiper-container {
    /*overflow: hidden;*/
    position: relative;
    width: 100%;
    margin: 0 auto;
    z-index: 1;
}
.swiper-wrapper {
    position: relative;
    width: 100%;
    z-index: 1;
}
</style>
						<div class="gnbWrap">
								<ul class="gnb">
								<%For conPLoop = 0 To ubound(conParentMenuName)
								conCurrent =""
										if menupos <> "" then  '-- �޴���ȣ �������� ����޴� �����ش�. 
											if ubound(split(menupos,"^"))>0 then
												if Cstr(split(menupos,"^")(0))  = Cstr(conPLoop) then
														conCurrent = " current"
												end if
											end if
										end if
								
								%>
								<% IF (conIsOffShopOpen and conPLoop=7) or conPLoop <> 7 THEN%>
								<li class="gnb0<%=conPLoop%> <%=conCurrent%>"><p><%=conParentMenuName(conPLoop)%></p><!-- <span><em></em></span> -->
									<div class="subNavi">
										<ul>
											<%
											For conLoop = 0 To conChildMenuSize(conPLoop)

											' ���õ� �ش� �Ŵ��� ���̺��� ���� menupos(idx) �� �����´�.
											if menupos<>"" then
												if isarray(split(menupos,"^")) then
													IF Cstr(split(menupos,"^")(1)) = Cstr(conLoop) THEN
														menuposidx = conMenuID(split(menupos,"^")(0),conLoop)
													end if
												end if
											end if
											%>
											<li>
												<% if instr(Replace(conChildMenuLinkUrl(conPLoop,conLoop),"/lectureadmin/","/diyadmin/"), "?") > 0 then %>
													<a href="<%=Replace(conChildMenuLinkUrl(conPLoop,conLoop),"/lectureadmin/","/diyadmin/")%>&menupos=<%=conPLoop%>^<%=conLoop%>">
												<% else %>
													<a href="<%=Replace(conChildMenuLinkUrl(conPLoop,conLoop),"/lectureadmin/","/diyadmin/")%>?menupos=<%=conPLoop%>^<%=conLoop%>">
												<% end if %>

												<%=conChildMenuName(conPLoop,conLoop)%></a>
											</li>
											<%Next%>
										</ul>
									</div>
								</li>
								<% END IF %>
								<%Next%>
								<!--<li class="swiper-slide gnbBtn"><a href="http://scm.10x10.co.kr/designer/index.asp" target="_blank"><img src="/images/partner/partner_btn_oldver.png" alt="�������� �ٷΰ���" /></a></li>-->
							</ul>
						 <div class="swiper-scrollbar"></div>
						</div>
							<script type="text/javascript" src="/js/jquery_common.js"></script>

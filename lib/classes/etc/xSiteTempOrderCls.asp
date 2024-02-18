<%
'###########################################################
' Description : 제휴몰 클래스
' Hieditor : 2011.04.22 이상구 생성
'			 2012.08.24 한용민 수정
'###########################################################

'/제휴상품명과 제휴옵션명으로 매칭하는 제휴몰		'/2013.04.24 한용민 생성
function GetItemMaeching_itemname_itemoptionname(sitename)
	if sitename = "" then
		GetItemMaeching_itemname_itemoptionname = FALSE
		exit function
	end if

	if sitename="bandinlunis" or sitename="mintstore" or sitename="byulshopITS" or sitename="itsByulshop"  or sitename="wconcept" or sitename="itsWconcept" or sitename="player" or sitename="itsPlayer1" or sitename="GVG" or sitename="gabangpop" or sitename="itsGabangpop" or sitename="musinsaITS" or sitename="itsMusinsa" or sitename="11stITS" or sitename="privia" or sitename="giftting" or sitename="ticketmonster" or sitename="coupang" or sitename="thinkaboutyou" or sitename="cookatmall" then
		GetItemMaeching_itemname_itemoptionname = TRUE
	else
		GetItemMaeching_itemname_itemoptionname = FALSE
	end if
end function

'/제휴상품명과 제휴옵션명으로 매칭하는 제휴몰		'/2013.04.24 한용민 생성
function GetItemMaeching_itemname_itemoptionname_list()
	response.write "* 반디앤루이스(bandinlunis)"
	response.write "&nbsp;&nbsp;* 민트샵(mintstore)"
	response.write "&nbsp;&nbsp;* 별샵(byulshopITS)"
	response.write "&nbsp;&nbsp;* 더블류컨셉(wconcept)"
	response.write "<br>* 플레이어(player)"
	response.write "&nbsp;&nbsp;* GVG(GVG)"
	response.write "&nbsp;&nbsp;* 가방팝(gabangpop)"
	response.write "&nbsp;&nbsp;* 무신사(musinsaITS)"
	response.write "&nbsp;&nbsp;* 11번가_아이띵소(11stITS)"
'	response.write "&nbsp;&nbsp;* gseshop(gseshop)"
	response.write "&nbsp;&nbsp;* 프리비아(privia)"
	response.write "&nbsp;&nbsp;* 아이띵소_29cm(its29cm)"
	response.write "&nbsp;&nbsp;* 기프팅(giftting)"
	response.write "&nbsp;&nbsp;* GS아이슈퍼(gsisuper)"
	response.write "&nbsp;&nbsp;* SUHA(suhaITS)"
	response.write "&nbsp;&nbsp;* 티켓몬스터(ticketmonster)"
	response.write "&nbsp;&nbsp;* 카카오기프트(kakaogift)"
	response.write "&nbsp;&nbsp;* 아이띵소샵(ithinksoshop)"
	response.write "&nbsp;&nbsp;* 띵크어바웃유(thinkaboutyou)"
	response.write "&nbsp;&nbsp;* 쿠캣(cookatmall)"
	response.write "&nbsp;&nbsp;* 맘큐(momQ)"
	response.write "&nbsp;&nbsp;* 미미박스(itsMemebox)"
	response.write "&nbsp;&nbsp;* 메이커스위드카카오(itsKaKaoMakers)"
end function

'/엑셀 주문 업로드 샘플
function get_xsite_excel_order_sample()		'/2013.04.24 한용민 생성
	''response.write "<a href='http://imgstatic.10x10.co.kr/offshop/sample/xsite/cn10x10_order_sample.xls' onfocus='this.blur()'>"
	''response.write "* 해외중국사이트(cn10x10)</a>"
	''response.write "<br><a href='http://imgstatic.10x10.co.kr/offshop/sample/xsite/mintstore_order_sample.xls' onfocus='this.blur()'>"
	''response.write "* 민트샵(mintstore)</a>"
	response.write "&nbsp;&nbsp;<a href='http://imgstatic.10x10.co.kr/offshop/sample/xsite/wconcept_order_sample.xls' onfocus='this.blur()'>"
	response.write "* 더블유컨셉(wconcept)</a>"
	response.write "&nbsp;&nbsp;<a href='http://imgstatic.10x10.co.kr/offshop/sample/xsite/byulshop_order_sample.xls' onfocus='this.blur()'>"
	response.write "* 별샵(byulshopITS)</a>"
	response.write "<br><a href='http://imgstatic.10x10.co.kr/offshop/sample/xsite/player_order_sample.xls' onfocus='this.blur()'>"
	response.write "* 플레이어(player)</a>"
	response.write "&nbsp;&nbsp;<a href='http://imgstatic.10x10.co.kr/offshop/sample/xsite/GVG_order_sample.xls' onfocus='this.blur()'>"
	response.write "* GVG(GVG)</a>"
	response.write "&nbsp;&nbsp;<a href='http://imgstatic.10x10.co.kr/offshop/sample/xsite/wizwid_order_sample.xls' onfocus='this.blur()'>"
	response.write "* 위즈위드(wizwid)</a>"
	response.write "&nbsp;&nbsp;<a href='http://imgstatic.10x10.co.kr/offshop/sample/xsite/hiphoper_order_sample.xls' onfocus='this.blur()'>"
	response.write "* 힙합퍼(hiphoper)</a>"
	response.write "<br><a href='http://imgstatic.10x10.co.kr/offshop/sample/xsite/gabangpop_order_sample.xls' onfocus='this.blur()'>"
	response.write "* 가방팝(gabangpop)</a>"
	response.write "&nbsp;&nbsp;<a href='http://imgstatic.10x10.co.kr/offshop/sample/xsite/musinsaITS_order_sample.xls' onfocus='this.blur()'>"
	response.write "* 무신사(musinsaITS)</a>"
	response.write "&nbsp;&nbsp;<a href='http://imgstatic.10x10.co.kr/offshop/sample/xsite/NJOYNY_order_sample.xls' onfocus='this.blur()'>"
	response.write "* 엔조이뉴욕(NJOYNY)</a>"
	response.write "<br><a href='http://imgstatic.10x10.co.kr/offshop/sample/xsite/fashionplus_order_sample.xls' onfocus='this.blur()'>"
	response.write "* 패션플러스(fashionplus)</a>"
	response.write "&nbsp;&nbsp;<a href='http://imgstatic.10x10.co.kr/offshop/sample/xsite/11stITS_order_sample.xls' onfocus='this.blur()'>"
	response.write "* 11번가_아이띵소(11stITS)</a>"
	response.write "&nbsp;&nbsp;<a href='http://imgstatic.10x10.co.kr/offshop/sample/xsite/GSeshop_order_sample.xls' onfocus='this.blur()'>"
	response.write "* GS SHOP(gseshop)</a>"
	response.write "<br><a href='http://imgstatic.10x10.co.kr/offshop/sample/xsite/Homeplus_order_sample.xls' onfocus='this.blur()'>"
	response.write "* Homeplus(homeplus)</a>"
	response.write "&nbsp;&nbsp;<a href='http://imgstatic.10x10.co.kr/offshop/sample/xsite/Ezwel_order_sample.xls' onfocus='this.blur()'>"
	response.write "* 이지웰페어(ezwel)</a>"
	response.write "<br><a href='http://imgstatic.10x10.co.kr/offshop/sample/xsite/kakaogift_order_sample.xls' onfocus='this.blur()'>"
	response.write "* 카카오기프트(kakaogift)</a>"
	response.write "&nbsp;&nbsp;<a href='http://imgstatic.10x10.co.kr/offshop/sample/xsite/hottracks_order_sample.xls' onfocus='this.blur()'>"
	response.write "* 교보 핫트랙스(hottracks)</a>"
	response.write "&nbsp;&nbsp;<a href='http://imgstatic.10x10.co.kr/offshop/sample/xsite/privia_order_sample2.xls' onfocus='this.blur()'>"
	response.write "* 프리비아(privia)</a>"
	response.write "<br><a href='http://imgstatic.10x10.co.kr/offshop/sample/xsite/dnshop_order_sample.xls' onfocus='this.blur()'>"
	response.write "* 디앤샵(dnshop)</a>"
	response.write "&nbsp;&nbsp;<a href='http://imgstatic.10x10.co.kr/offshop/sample/xsite/its29cm_order_sample.xls' onfocus='this.blur()'>"
	response.write "* 아이띵소_29cm(its29cm)</a>"
	response.write "&nbsp;&nbsp;<a href='http://imgstatic.10x10.co.kr/offshop/sample/xsite/suhaITS_order_sample.xls' onfocus='this.blur()'>"
	response.write "* 아이띵소_SUHA(suhaITS)</a>"
	response.write "<br><a href='http://imgstatic.10x10.co.kr/offshop/sample/xsite/ticketmonster_order_sample.xls' onfocus='this.blur()'>"
	response.write "* 티켓몬스터(ticketmonster)</a>"
	response.write "&nbsp;&nbsp;<a href='http://imgstatic.10x10.co.kr/offshop/sample/xsite/itsCjmall_order_sample.xls' onfocus='this.blur()'>"
	response.write "* 아이띵소_cjmall(itsCjmall)</a>"
	response.write "&nbsp;&nbsp;<a href='http://imgstatic.10x10.co.kr/offshop/sample/xsite/ithinksoshop_order_sample2.xls' onfocus='this.blur()'>"
	response.write "<br>* 아이띵소샵(ithinksoshop)</a>"
	response.write "&nbsp;&nbsp;<a href='http://imgstatic.10x10.co.kr/offshop/sample/xsite/momastore_order_sample.xls' onfocus='this.blur()'>"
	response.write "* momastore(momastore)</a>"
	response.write "&nbsp;&nbsp;<a href='http://imgstatic.10x10.co.kr/offshop/sample/xsite/kostkamm_order_sample.xls' onfocus='this.blur()'>"
	response.write "<br>* 띵크어바웃유(thinkaboutyou)"
	response.write "&nbsp;&nbsp;<a href='http://imgstatic.10x10.co.kr/offshop/sample/xsite/momQ_order_sample.xls' onfocus='this.blur()'>"
	response.write "* 맘큐(momQ)"
	response.write "&nbsp;&nbsp;<a href='http://imgstatic.10x10.co.kr/offshop/sample/xsite/itsMemebox_order_sample.xls' onfocus='this.blur()'>"
	response.write "* 미미박스(itsMemebox)"
	response.write "<br><a href='http://imgstatic.10x10.co.kr/offshop/sample/xsite/cnglob10x10_order_sample.xls' onfocus='this.blur()'>"
	response.write "* 해외중국사이트(cnglob10x10)"
	response.write "&nbsp;&nbsp;<a href='http://imgstatic.10x10.co.kr/offshop/sample/xsite/cnhigo_order_sample.xls' onfocus='this.blur()'>"
	response.write "* 해외HIGO(cnhigo)"
	response.write "&nbsp;&nbsp;<a href='http://imgstatic.10x10.co.kr/offshop/sample/xsite/itsKaKaoMakers_order_sample.xls' onfocus='this.blur()'>"
	response.write "* 메이커스위드카카오(itsKaKaoMakers)"
end function

''상품 수기 관리 제휴사
public Sub drawSelectBoxXSiteHandItemPartner(selBoxName, selVal)
    dim retStr
    retStr = "<select name='"&selBoxName&"'>"
	retStr = retStr & " <option value=''>선택"
	retStr = retStr & " <option value='interparkPTM' "& CHKIIF(selVal="interparkPTM","selected","") &" >인터파크프레즌투미"
	''retStr = retStr & " <option value='mintstore' "& CHKIIF(selVal="mintstore","selected","") &" >민트샵"
	retStr = retStr & " <option value='fashionplus' "& CHKIIF(selVal="fashionplus","selected","") &" >패션플러스"
	retStr = retStr & " <option value='cjmall' "& CHKIIF(selVal="cjmall","selected","") &" >10x10_씨제이(아이띵소x)"
	retStr = retStr & " <option value='wizwid' "& CHKIIF(selVal="wizwid","selected","") &" >위즈위드"
	retStr = retStr & " <option value='wconcept' "& CHKIIF(selVal="wconcept","selected","") &" >더블유컨셉"
	''retStr = retStr & " <option value='ollehtv' "& CHKIIF(selVal="ollehtv","selected","") &" >올레TV"
	retStr = retStr & " <option value='hottracks' "& CHKIIF(selVal="hottracks","selected","") &" >교보핫트랙스"
	retStr = retStr & " <option value='hiphoper' "& CHKIIF(selVal="hiphoper","selected","") &" >힙합퍼"
	''retStr = retStr & " <option value='byulshopITS' "& CHKIIF(selVal="byulshopITS","selected","") &" >별샵"
	retStr = retStr & " <option value='cjmallITS' "& CHKIIF(selVal="cjmallITS","selected","") &" >cjmallITS" ''(아이띵소)
	retStr = retStr & " <option value='player' "& CHKIIF(selVal="player","selected","") &" >player"
	retStr = retStr & " <option value='gabangpop' "& CHKIIF(selVal="gabangpop","selected","") &" >가방팝"
	'retStr = retStr & " <option value='musinsaITS' "& CHKIIF(selVal="musinsaITS","selected","") &" >무신사"
	'retStr = retStr & " <option value='NJOYNY' "& CHKIIF(selVal="NJOYNY","selected","") &" >엔조이뉴욕"
	retStr = retStr & " <option value='gseshop' "& CHKIIF(selVal="gseshop","selected","") &" >gseshop"
	retStr = retStr & " <option value='homeplus' "& CHKIIF(selVal="homeplus","selected","") &" >Homeplus"
	retStr = retStr & " <option value='ezwel' "& CHKIIF(selVal="ezwel","selected","") &" >이지웰페어"
	retStr = retStr & " <option value='benepia1010' "& CHKIIF(selVal="benepia1010","selected","") &" >베네피아"
	retStr = retStr & " <option value='boribori1010' "& CHKIIF(selVal="boribori1010","selected","") &" >보리보리"
	retStr = retStr & " <option value='bindmall1010' "& CHKIIF(selVal="bindmall1010","selected","") &" >바인드웍스"
	retStr = retStr & " <option value='GS25' "& CHKIIF(selVal="GS25","selected","") &" >GS25카달로그"
	retStr = retStr & " <option value='privia' "& CHKIIF(selVal="privia","selected","") &" >프리비아"
	retStr = retStr & " <option value='momastore' "& CHKIIF(selVal="momastore","selected","") &" >momastore"
	retStr = retStr & " <option value='bandinlunis' "& CHKIIF(selVal="bandinlunis","selected","") &" >반디앤루이스"
	retStr = retStr & " <option value='coupang' "& CHKIIF(selVal="coupang","selected","") &" >쿠팡"
	retStr = retStr & " <option value='hmall1010' "& CHKIIF(selVal="hmall1010","selected","") &" >Hmall"
	retStr = retStr & " <option value='giftting' "& CHKIIF(selVal="giftting","selected","") &" >기프팅"
	retStr = retStr & " <option value='kakaogift' "& CHKIIF(selVal="kakaogift","selected","") &" >카카오기프트"
	retStr = retStr & " <option value='kakaostore' "& CHKIIF(selVal="kakaostore","selected","") &" >카카오톡스토어"
	retStr = retStr & " <option value='cookatmall' "& CHKIIF(selVal="cookatmall","selected","") &" >쿠캣"
	retStr = retStr & " <option value='ticketmonster' "& CHKIIF(selVal="ticketmonster","selected","") &" >티켓몬스터"
	retStr = retStr & " <option value='thinkaboutyou' "& CHKIIF(selVal="thinkaboutyou","selected","") &" >띵크어바웃유"
	retStr = retStr & " <option value='momQ' "& CHKIIF(selVal="momQ","selected","") &" >맘큐"
	retStr = retStr & " <option value='gmarket' "& CHKIIF(selVal="gmarket","selected","") &" >gmarket"
	retStr = retStr & " <option value='celectory' "& CHKIIF(selVal="celectory","selected","") &" >셀렉토리"
	retStr = retStr & " <option value='gsisuper' "& CHKIIF(selVal="gsisuper","selected","") &" >GS아이슈퍼"
	retStr = retStr & " <option value='etsy' "& CHKIIF(selVal="etsy","selected","") &" >[해외etsy]"
	retStr = retStr & " <option value='' >---------------------"
	retStr = retStr & " <option value='itsByulshop' "& CHKIIF(selVal="itsByulshop","selected","") &" >아이띵소_별샵"
	retStr = retStr & " <option value='itsCjmall' "& CHKIIF(selVal="itsCjmall","selected","") &" >아이띵소_cjmall"
	retStr = retStr & " <option value='itsFashionplus' "& CHKIIF(selVal="itsFashionplus","selected","") &" >아이띵소_패션플러스"
	retStr = retStr & " <option value='itsGabangpop' "& CHKIIF(selVal="itsGabangpop","selected","") &" >아이띵소_가방팝"
	retStr = retStr & " <option value='itsHiphoper' "& CHKIIF(selVal="itsHiphoper","selected","") &" >아이띵소_힙합퍼"
	retStr = retStr & " <option value='itsHottracks' "& CHKIIF(selVal="itsHottracks","selected","") &" >아이띵소_교보핫트랙스"
	''retStr = retStr & " <option value='itsMintstore' "& CHKIIF(selVal="itsMintstore","selected","") &" >아이띵소_민트샵"
	retStr = retStr & " <option value='itsMusinsa' "& CHKIIF(selVal="itsMusinsa","selected","") &" >아이띵소_무신사"
	retStr = retStr & " <option value='itsNJOYNY' "& CHKIIF(selVal="itsNJOYNY","selected","") &" >아이띵소_엔조이뉴욕"
	retStr = retStr & " <option value='itsPlayer1' "& CHKIIF(selVal="itsPlayer1","selected","") &" >아이띵소_player"
	retStr = retStr & " <option value='itsWconcept' "& CHKIIF(selVal="itsWconcept","selected","") &" >아이띵소_더블유컨셉"
	retStr = retStr & " <option value='itsWizwid' "& CHKIIF(selVal="itsWizwid","selected","") &" >아이띵소_위즈위드"
    retStr = retStr & " <option value='its29cm' "& CHKIIF(selVal="its29cm","selected","") &" >아이띵소_29cm"
    retStr = retStr & " <option value='suhaITS' "& CHKIIF(selVal="suhaITS","selected","") &" >아이띵소_SUHA"
    retStr = retStr & " <option value='ithinksoshop' "& CHKIIF(selVal="ithinksoshop","selected","") &" >아이띵소샵"
    retStr = retStr & " <option value='itsMemebox' "& CHKIIF(selVal="itsMemebox","selected","") &" >아이띵소_미미박스"
    retStr = retStr & " <option value='itsKaKaoMakers' "& CHKIIF(selVal="itsKaKaoMakers","selected","") &" >아이띵소_메이커스위드카카오"
    retStr = retStr & " <option value=' itskakao' "& CHKIIF(selVal="itskakao","selected","") &" >아이띵소_카카오선물하기"
    retStr = retStr & " <option value='itsWadiz' "& CHKIIF(selVal="itsWadiz","selected","") &" >아이띵소_와디즈"
	retStr = retStr & " <option value='itsbenepia' "& CHKIIF(selVal="itsbenepia","selected","") &" >아이띵소_베네피아"
	retStr = retStr & " <option value='itskakaotalkstore' "& CHKIIF(selVal="itskakaotalkstore","selected","") &" >아이띵소_카카오톡스토어"
	''퇴점사
	''retStr = retStr & " <option value='cn10x10' "& CHKIIF(selVal="cn10x10","selected","") &" >해외중국사이트"
	''retStr = retStr & " <option value='11stITS' "& CHKIIF(selVal="11stITS","selected","") &" >11번가_아이띵소"
	''retStr = retStr & " <option value='GVG' "& CHKIIF(selVal="GVG","selected","") &" >GVG"
	''retStr = retStr & " <option value='gmarket' "& CHKIIF(selVal="gmarket","selected","") &" >gmarket"
	''retStr = retStr & " <option value='hanatour' "& CHKIIF(selVal="hanatour","selected","") &" >하나투어"
	retStr = retStr & " </select> "

	response.write retStr
end Sub

public Sub drawSelectBoxXSiteOrderInputPartnerCS(selBoxName, selVal)
	dim retStr
	retStr = "<select class='select' name='"&selBoxName&"'>" & vbCrLf
	retStr = retStr & " <option value=''>선택</option>" & vbCrLf
	retStr = retStr & " <option value='10x10' "& CHKIIF(selVal="10x10","selected","") &" >10x10</option>" & vbCrLf
	retStr = retStr & " <option value='extall' "& CHKIIF(selVal="extall","selected","") &" >제휴몰전체</option>" & vbCrLf
	retStr = retStr & " <option value='' >---------------------</option>" & vbCrLf
	retStr = retStr & " <option value='11st1010' "& CHKIIF(selVal="11st1010","selected","") &" >11번가"
	retStr = retStr & " <option value='cjmall' "& CHKIIF(selVal="cjmall","selected","") &" >cjMall</option>" & vbCrLf
	retStr = retStr & " <option value='lotteimall' "& CHKIIF(selVal="lotteimall","selected","") &" >롯데iMall</option>" & vbCrLf
	retStr = retStr & " <option value='ssg' "& CHKIIF(selVal="ssg","selected","") &" >신세계몰(SSG)</option>" & vbCrLf
	' retStr = retStr & " <option value='lotteCom' "& CHKIIF(selVal="lotteCom","selected","") &" >롯데닷컴</option>"
	retStr = retStr & " <option value='lotteon' "& CHKIIF(selVal="lotteon","selected","") &" >롯데On</option>"
	retStr = retStr & " <option value='shintvshopping' "& CHKIIF(selVal="shintvshopping","selected","") &" >신세계TV쇼핑</option>"
	retStr = retStr & " <option value='skstoa' "& CHKIIF(selVal="skstoa","selected","") &" >SKSTOA</option>"
	retStr = retStr & " <option value='wetoo1300k' "& CHKIIF(selVal="wetoo1300k","selected","") &" >1300k</option>"
	retStr = retStr & " <option value='gseshop' "& CHKIIF(selVal="gseshop","selected","") &" >gseshop</option>" & vbCrLf
	retStr = retStr & " <option value='hmall1010' "& CHKIIF(selVal="hmall1010","selected","") &" >HMall</option>" & vbCrLf
	retStr = retStr & " <option value='ezwel' "& CHKIIF(selVal="ezwel","selected","") &" >이지웰페어</option>" & vbCrLf
	retStr = retStr & " <option value='benepia1010' "& CHKIIF(selVal="benepia1010","selected","") &" >베네피아</option>" & vbCrLf
	retStr = retStr & " <option value='boribori1010' "& CHKIIF(selVal="boribori1010","selected","") &" >보리보리</option>" & vbCrLf
	retStr = retStr & " <option value='bindmall1010' "& CHKIIF(selVal="bindmall1010","selected","") &" >바인드웍스</option>" & vbCrLf
'	retStr = retStr & " <option value='halfclub' "& CHKIIF(selVal="halfclub","selected","") &" >하프클럽</option>" & vbCrLf
	retStr = retStr & " <option value='nvstorefarm' "& CHKIIF(selVal="nvstorefarm","selected","") &" >스토어팜</option>" & vbCrLf
'	retStr = retStr & " <option value='nvstoremoonbangu' "& CHKIIF(selVal="nvstoremoonbangu","selected","") &" >스토어팜 문방구</option>" & vbCrLf
	retStr = retStr & " <option value='Mylittlewhoopee' "& CHKIIF(selVal="Mylittlewhoopee","selected","") &" >스토어팜 캣앤독</option>" & vbCrLf
	retStr = retStr & " <option value='nvstoregift' "& CHKIIF(selVal="nvstoregift","selected","") &" >스토어팜 선물하기</option>" & vbCrLf
	retStr = retStr & " <option value='auction1010' "& CHKIIF(selVal="auction1010","selected","") &" >옥션</option>" & vbCrLf
	retStr = retStr & " <option value='gmarket1010' "& CHKIIF(selVal="gmarket1010","selected","") &" >지마켓(New)</option>" & vbCrLf
	retStr = retStr & " <option value='interpark' "& CHKIIF(selVal="interpark","selected","") &" >인터파크</option>" & vbCrLf
	retStr = retStr & " <option value='coupang' "& CHKIIF(selVal="coupang","selected","") &" >쿠팡</option>" & vbCrLf
'	retStr = retStr & " <option value='wemakeprice' "& CHKIIF(selVal="wemakeprice","selected","") &" >위메프</option>" & vbCrLf
	retStr = retStr & " <option value='WMP' "& CHKIIF(selVal="WMP","selected","") &" >위메프(API)</option>"
'	retStr = retStr & " <option value='wmpfashion' "& CHKIIF(selVal="wmpfashion","selected","") &" >위메프W패션(API)</option>"
	' retStr = retStr & " <option value='ssgmemo' "& CHKIIF(selVal="ssgmemo","selected","") &" >신세계몰(SSG 메모)</option>" & vbCrLf
	retStr = retStr & " <option value='lfmall' "& CHKIIF(selVal="lfmall","selected","") &" >LFmall</option>" & vbCrLf
	retStr = retStr & " <option value='wconcept1010' "& CHKIIF(selVal="wconcept1010","selected","") &" >W컨셉</option>" & vbCrLf
	retStr = retStr & " <option value='withnature1010' "& CHKIIF(selVal="withnature1010","selected","") &" >자연이랑</option>" & vbCrLf
	retStr = retStr & " <option value='goodshop1010' "& CHKIIF(selVal="goodshop1010","selected","") &" >굿샵</option>" & vbCrLf
	' retStr = retStr & " <option value='yes24' "& CHKIIF(selVal="yes24","selected","") &" >YES24</option>" & vbCrLf
	' retStr = retStr & " <option value='alphamall' "& CHKIIF(selVal="alphamall","selected","") &" >알파몰"
	' retStr = retStr & " <option value='ohou1010' "& CHKIIF(selVal="ohou1010","selected","") &" >오늘의집"
    '	retStr = retStr & " <option value='wadsmartstore' "& CHKIIF(selVal="wadsmartstore","selected","") &" >와드스마트스토어"
    retStr = retStr & " <option value='casamia_good_com' "& CHKIIF(selVal="casamia_good_com","selected","") &" >까사미아"
    retStr = retStr & " <option value='kakaogift' "& CHKIIF(selVal="kakaogift","selected","") &" >카카오기프트"
	retStr = retStr & " <option value='kakaostore' "& CHKIIF(selVal="kakaostore","selected","") &" >카카오톡스토어"
	retStr = retStr & " <option value='cookatmall' "& CHKIIF(selVal="cookatmall","selected","") &" >쿠캣"
    retStr = retStr & " <option value='alphamall' "& CHKIIF(selVal="alphamall","selected","") &" >알파몰"
    retStr = retStr & " <option value='aboutpet' "& CHKIIF(selVal="aboutpet","selected","") &" >어바웃펫"
	retStr = retStr & " <option value='goodwearmall10' "& CHKIIF(selVal="goodwearmall10","selected","") &" >굿웨어몰"
	retStr = retStr & " <option value='shopify' "& CHKIIF(selVal="shopify","selected","") &" >shopify"
	retStr = retStr & " </select> "
	response.write retStr
end Sub

public Sub drawSelectBoxXSiteOrderInputPartner(selBoxName, selVal)
    dim retStr
    retStr = "<select class='select' name='"&selBoxName&"'>"
	retStr = retStr & " <option value=''>선택"
	retStr = retStr & " <option value='ssg' "& CHKIIF(selVal="ssg","selected","") &" >신세계몰(SSG)"
	retStr = retStr & " <option value='interpark' "& CHKIIF(selVal="interpark","selected","") &" >인터파크"
	retStr = retStr & " <option value='cjmall' "& CHKIIF(selVal="cjmall","selected","") &" >cjMall"
	retStr = retStr & " <option value='coupang' "& CHKIIF(selVal="coupang","selected","") &" >쿠팡"
	retStr = retStr & " <option value='11st1010' "& CHKIIF(selVal="11st1010","selected","") &" >11번가"
	retStr = retStr & " <option value='ezwel' "& CHKIIF(selVal="ezwel","selected","") &" >이지웰페어"
	retStr = retStr & " <option value='benepia1010' "& CHKIIF(selVal="benepia1010","selected","") &" >베네피아</option>"
	retStr = retStr & " <option value='boribori1010' "& CHKIIF(selVal="boribori1010","selected","") &" >보리보리"
	retStr = retStr & " <option value='bindmall1010' "& CHKIIF(selVal="bindmall1010","selected","") &" >바인드웍스"
	retStr = retStr & " <option value='gmarket1010' "& CHKIIF(selVal="gmarket1010","selected","") &" >지마켓(New)"
	retStr = retStr & " <option value='gseshop' "& CHKIIF(selVal="gseshop","selected","") &" >gseshop"
'	retStr = retStr & " <option value='lotteCom' "& CHKIIF(selVal="lotteCom","selected","") &" >롯데닷컴"
	retStr = retStr & " <option value='lotteimall' "& CHKIIF(selVal="lotteimall","selected","") &" >롯데iMall"
	retStr = retStr & " <option value='lotteon' "& CHKIIF(selVal="lotteon","selected","") &" >롯데On"
	retStr = retStr & " <option value='shintvshopping' "& CHKIIF(selVal="shintvshopping","selected","") &" >신세계TV쇼핑</option>"
	retStr = retStr & " <option value='skstoa' "& CHKIIF(selVal="skstoa","selected","") &" >SKSTOA</option>"
	retStr = retStr & " <option value='wetoo1300k' "& CHKIIF(selVal="wetoo1300k","selected","") &" >1300k</option>"
	retStr = retStr & " <option value='nvstorefarm' "& CHKIIF(selVal="nvstorefarm","selected","") &" >스토어팜"
'	retStr = retStr & " <option value='nvstoremoonbangu' "& CHKIIF(selVal="nvstoremoonbangu","selected","") &" >스토어팜 문방구"
	retStr = retStr & " <option value='Mylittlewhoopee' "& CHKIIF(selVal="Mylittlewhoopee","selected","") &" >스토어팜 캣앤독"
	retStr = retStr & " <option value='nvstoregift' "& CHKIIF(selVal="nvstoregift","selected","") &" >스토어팜 선물하기"
	retStr = retStr & " <option value='auction1010' "& CHKIIF(selVal="auction1010","selected","") &" >옥션"
	retStr = retStr & " <option value='hmall1010' "& CHKIIF(selVal="hmall1010","selected","") &" >HMall"
	retStr = retStr & " <option value='wemakeprice' "& CHKIIF(selVal="wemakeprice","selected","") &" >위메프"
	retStr = retStr & " <option value='WMP' "& CHKIIF(selVal="WMP","selected","") &" >위메프(API)"
	retStr = retStr & " <option value='wmpfashion' "& CHKIIF(selVal="wmpfashion","selected","") &" >위메프W패션(API)"
	retStr = retStr & " <option value='kakaogift' "& CHKIIF(selVal="kakaogift","selected","") &" >카카오기프트"
	retStr = retStr & " <option value='kakaostore' "& CHKIIF(selVal="kakaostore","selected","") &" >카카오톡스토어"
	retStr = retStr & " <option value='cookatmall' "& CHKIIF(selVal="cookatmall","selected","") &" >쿠캣"
'	retStr = retStr & " <option value='nvstorefarmclass' "& CHKIIF(selVal="nvstorefarmclass","selected","") &" >스토어팜클래스"
	retStr = retStr & " <option value='giftting' "& CHKIIF(selVal="giftting","selected","") &" >기프팅"
'	retStr = retStr & " <option value='halfclub' "& CHKIIF(selVal="halfclub","selected","") &" >하프클럽"
'	retStr = retStr & " <option value='gsisuper' "& CHKIIF(selVal="gsisuper","selected","") &" >GS아이슈퍼"
	retStr = retStr & " <option value='LFmall' "& CHKIIF(selVal="LFmall","selected","") &" >LFmall"
	retStr = retStr & " <option value='wconcept1010' "& CHKIIF(selVal="wconcept1010","selected","") &" >W컨셉"
	retStr = retStr & " <option value='withnature1010' "& CHKIIF(selVal="withnature1010","selected","") &" >자연이랑"
	retStr = retStr & " <option value='goodshop1010' "& CHKIIF(selVal="goodshop1010","selected","") &" >굿샵</option>" & vbCrLf
	retStr = retStr & " <option value='yes24' "& CHKIIF(selVal="yes24","selected","") &" >YES24"
	retStr = retStr & " <option value='alphamall' "& CHKIIF(selVal="alphamall","selected","") &" >알파몰"
	retStr = retStr & " <option value='ohou1010' "& CHKIIF(selVal="ohou1010","selected","") &" >오늘의집"
	retStr = retStr & " <option value='wadsmartstore' "& CHKIIF(selVal="wadsmartstore","selected","") &" >와드스마트스토어"
	retStr = retStr & " <option value='casamia_good_com' "& CHKIIF(selVal="casamia_good_com","selected","") &" >까사미아"
	retStr = retStr & " <option value='aboutpet' "& CHKIIF(selVal="aboutpet","selected","") &" >어바웃펫"
	retStr = retStr & " <option value='goodwearmall10' "& CHKIIF(selVal="goodwearmall10","selected","") &" >굿웨어몰"
	retStr = retStr & " <option value='GS25' "& CHKIIF(selVal="GS25","selected","") &" >GS25카달로그"
	retStr = retStr & " <option value='shopify' "& CHKIIF(selVal="shopify","selected","") &" >shopify"
	retStr = retStr & " <option value='shoplinker' "& CHKIIF(selVal="shoplinker","selected","") &" >샵링커###########"
	''retStr = retStr & " <option value='homeplus' "& CHKIIF(selVal="homeplus","selected","") &" >Homeplus"
	'retStr = retStr & " <option value='cn10x10' "& CHKIIF(selVal="cn10x10","selected","") &" >해외중국사이트"
	''retStr = retStr & " <option value='cnglob10x10' "& CHKIIF(selVal="cnglob10x10","selected","") &" >[해외중국사이트]"
	''retStr = retStr & " <option value='cnhigo' "& CHKIIF(selVal="cnhigo","selected","") &" >[해외HIGO]"
	''retStr = retStr & " <option value='cnugoshop' "& CHKIIF(selVal="cnugoshop","selected","") &" >[해외UGOSHOP]"
	''retStr = retStr & " <option value='11stmy' "& CHKIIF(selVal="11stmy","selected","") &" >[해외11번가]"
	''retStr = retStr & " <option value='etsy' "& CHKIIF(selVal="etsy","selected","") &" >[해외etsy]"
	'retStr = retStr & " <option value='zilingo' "& CHKIIF(selVal="zilingo","selected","") &" >[해외Zilingo]"
	'retStr = retStr & " <option value='ticketmonster' "& CHKIIF(selVal="ticketmonster","selected","") &" >티켓몬스터"
	'retStr = retStr & " <option value='thinkaboutyou' "& CHKIIF(selVal="thinkaboutyou","selected","") &" >띵크어바웃유"
	'retStr = retStr & " <option value='momQ' "& CHKIIF(selVal="momQ","selected","") &" >맘큐"
	'retStr = retStr & " <option value='gmarket' "& CHKIIF(selVal="gmarket","selected","") &" >gmarket(OLD)"
	'retStr = retStr & " <option value='celectory' "& CHKIIF(selVal="celectory","selected","") &" >셀렉토리"

	retStr = retStr & " <option value='' >---------------------"
	'retStr = retStr & " <option value='lotteComM' "& CHKIIF(selVal="lotteComM","selected","") &" >롯데닷컴(매입)"
	''retStr = retStr & " <option value='mintstore' "& CHKIIF(selVal="mintstore","selected","") &" >민트샵"
	'retStr = retStr & " <option value='fashionplus' "& CHKIIF(selVal="fashionplus","selected","") &" >패션플러스"
	'retStr = retStr & " <option value='wizwid' "& CHKIIF(selVal="wizwid","selected","") &" >위즈위드"
	'retStr = retStr & " <option value='wconcept' "& CHKIIF(selVal="wconcept","selected","") &" >W컨셉"

	'retStr = retStr & " <option value='29cm' "& CHKIIF(selVal="29cm","selected","") &" >29cm"
	''retStr = retStr & " <option value='ollehtv' "& CHKIIF(selVal="ollehtv","selected","") &" >올레TV"
	'retStr = retStr & " <option value='hottracks' "& CHKIIF(selVal="hottracks","selected","") &" >교보핫트랙스"
	'retStr = retStr & " <option value='hiphoper' "& CHKIIF(selVal="hiphoper","selected","") &" >힙합퍼"
	''retStr = retStr & " <option value='byulshopITS' "& CHKIIF(selVal="byulshopITS","selected","") &" >별샵"
	'retStr = retStr & " <option value='cjmallITS' "& CHKIIF(selVal="cjmallITS","selected","") &" >cjmall아이띵소"
	'retStr = retStr & " <option value='player' "& CHKIIF(selVal="player","selected","") &" >player"
	'retStr = retStr & " <option value='gabangpop' "& CHKIIF(selVal="gabangpop","selected","") &" >가방팝"
	'retStr = retStr & " <option value='musinsaITS' "& CHKIIF(selVal="musinsaITS","selected","") &" >무신사"
	'retStr = retStr & " <option value='NJOYNY' "& CHKIIF(selVal="NJOYNY","selected","") &" >엔조이뉴욕"
'	retStr = retStr & " <option value='gseshop' "& CHKIIF(selVal="gseshop","selected","") &" >GS SHOP"
	'retStr = retStr & " <option value='privia' "& CHKIIF(selVal="privia","selected","") &" >프리비아"
	'retStr = retStr & " <option value='momastore' "& CHKIIF(selVal="momastore","selected","") &" >momastore"
	'retStr = retStr & " <option value='dnshop' "& CHKIIF(selVal="dnshop","selected","") &" >디앤샵"
	'retStr = retStr & " <option value='bandinlunis' "& CHKIIF(selVal="bandinlunis","selected","") &" >반디앤루이스"
	'retStr = retStr & " <option value='thinkaboutyou' "& CHKIIF(selVal="thinkaboutyou","selected","") &" >띵크어바웃유"

	retStr = retStr & " <option value='' >---------------------"
	retStr = retStr & " <option value='itsByulshop' "& CHKIIF(selVal="itsByulshop","selected","") &" >아이띵소_별샵"
	retStr = retStr & " <option value='itsCjmall' "& CHKIIF(selVal="itsCjmall","selected","") &" >아이띵소_cjmall"
	retStr = retStr & " <option value='itsFashionplus' "& CHKIIF(selVal="itsFashionplus","selected","") &" >아이띵소_패션플러스"
	retStr = retStr & " <option value='itsGabangpop' "& CHKIIF(selVal="itsGabangpop","selected","") &" >아이띵소_가방팝"
	retStr = retStr & " <option value='itsHiphoper' "& CHKIIF(selVal="itsHiphoper","selected","") &" >아이띵소_힙합퍼"
	retStr = retStr & " <option value='itsHottracks' "& CHKIIF(selVal="itsHottracks","selected","") &" >아이띵소_교보핫트랙스"
	''retStr = retStr & " <option value='itsMintstore' "& CHKIIF(selVal="itsMintstore","selected","") &" >아이띵소_민트샵"
	retStr = retStr & " <option value='itsMusinsa' "& CHKIIF(selVal="itsMusinsa","selected","") &" >아이띵소_무신사"
	retStr = retStr & " <option value='itsNJOYNY' "& CHKIIF(selVal="itsNJOYNY","selected","") &" >아이띵소_엔조이뉴욕"
	retStr = retStr & " <option value='itsPlayer1' "& CHKIIF(selVal="itsPlayer1","selected","") &" >아이띵소_player"
	retStr = retStr & " <option value='itsWconcept' "& CHKIIF(selVal="itsWconcept","selected","") &" >아이띵소_더블유컨셉"
	retStr = retStr & " <option value='itsWizwid' "& CHKIIF(selVal="itsWizwid","selected","") &" >아이띵소_위즈위드"
	retStr = retStr & " <option value='its29cm' "& CHKIIF(selVal="its29cm","selected","") &" >아이띵소_29cm"
	retStr = retStr & " <option value='suhaITS' "& CHKIIF(selVal="suhaITS","selected","") &" >아이띵소_SUHA"
	retStr = retStr & " <option value='ithinksoshop' "& CHKIIF(selVal="ithinksoshop","selected","") &" >아이띵소샵"
	retStr = retStr & " <option value='itsMemebox' "& CHKIIF(selVal="itsMemebox","selected","") &" >아이띵소_미미박스"
	retStr = retStr & " <option value='itsKaKaoMakers' "& CHKIIF(selVal="itsKaKaoMakers","selected","") &" >아이띵소_메이커스위드카카오"
	retStr = retStr & " <option value='itskakao' "& CHKIIF(selVal="itskakao","selected","") &" >아이띵소_카카오선물하기"
	retStr = retStr & " <option value='itsWadiz' "& CHKIIF(selVal="itsWadiz","selected","") &" >아이띵소_와디즈"
	retStr = retStr & " <option value='itsbenepia' "& CHKIIF(selVal="itsbenepia","selected","") &" >아이띵소_베네피아"
	retStr = retStr & " <option value='itskakaotalkstore' "& CHKIIF(selVal="itskakaotalkstore","selected","") &" >아이띵소_카카오톡스토어"

	'2013-11-25 김진영 추가// 샵링커로 주문입력 하는 부분이라 주문리스트에만 나오기 위함
	If Request.ServerVariables("Script_Name") = "/admin/ordermaster/outmalllist.asp" Then
		retStr = retStr & " <option value='itsGsshop' "& CHKIIF(selVal="itsGsshop","selected","") &" >아이띵소_GS SHOP"
		retStr = retStr & " <option value='itsDnshop' "& CHKIIF(selVal="itsDnshop","selected","") &" >아이띵소_디엔샵"
		retStr = retStr & " <option value='its11st' "& CHKIIF(selVal="its11st","selected","") &" >아이띵소_11번가"
		retStr = retStr & " <option value='itsGmarket' "& CHKIIF(selVal="itsGmarket","selected","") &" >아이띵소_G마켓"
		retStr = retStr & " <option value='itsShinsegae' "& CHKIIF(selVal="itsShinsegae","selected","") &" >아이띵소_신세계"
		retStr = retStr & " <option value='itsShinsegaeDept' "& CHKIIF(selVal="itsShinsegaeDept","selected","") &" >아이띵소_신세계백화점"
		retStr = retStr & " <option value='itssmartstore' "& CHKIIF(selVal="itssmartstore","selected","") &" >아이띵소_스마트스토어"
	End If
	''retStr = retStr & " <option value='11stITS' "& CHKIIF(selVal="11stITS","selected","") &" >11번가_아이띵소"
	''retStr = retStr & " <option value='GVG' "& CHKIIF(selVal="GVG","selected","") &" >GVG"
	''retStr = retStr & " <option value='gmarket' "& CHKIIF(selVal="gmarket","selected","") &" >gmarket"
	''retStr = retStr & " <option value='hanatour' "& CHKIIF(selVal="hanatour","selected","") &" >하나투어"
	retStr = retStr & " </select> "

	response.write retStr
end Sub

Public Function getIsValidItemIdOption(iitemid, iitemoption)
	Dim sqlStr, itemCount, itemOptionCount, itemOptionWithCount
	sqlStr = ""
	sqlStr = sqlStr & " SELECT Count(*) as cnt FROM db_item.dbo.tbl_item WHERE itemid = '"& iitemid &"' "
	rsget.Open sqlStr,dbget,1
		itemCount = rsget("cnt")
	rsget.Close
rw itemCount
	sqlStr = ""
	sqlStr = sqlStr & " SELECT Count(*) as cnt FROM db_item.dbo.tbl_item_option WHERE itemid = '"& iitemid &"' "
	rsget.Open sqlStr,dbget,1
		itemOptionCount = rsget("cnt")
	rsget.Close
rw itemOptionCount
	sqlStr = ""
	sqlStr = sqlStr & " SELECT Count(*) as cnt FROM db_item.dbo.tbl_item_option WHERE itemid = '"& iitemid &"' and itemoption = '"& iitemoption &"' "
	rsget.Open sqlStr,dbget,1
		itemOptionWithCount = rsget("cnt")
	rsget.Close
rw itemOptionWithCount
	If itemCount < 1 OR (itemOptionCount > 0 AND itemOptionWithCount < 1) Then
		getIsValidItemIdOption = "N"
	Else
		getIsValidItemIdOption = "Y"
	End If
End Function

public function getEtcSiteNameOrCode2ItemCode(byval sellsite, byval xsiteItemID, byval extitemname, byval extitemoptionname, byref rtitemid, byref rtitemoption, byref rtSellPrice)
    dim sqlStr, isTooMany

    sqlStr = " select top 10 T.* , i.itemname,i.sellyn"
    IF (application("Svr_Info")	= "Dev") then
        sqlStr = sqlStr & ",isNULL(i.sellcash,0) as sellcash"
        sqlStr = sqlStr & " from db_temp.dbo.tbl_xSite_EtcItemLink T"
        sqlStr = sqlStr & " Left join db_item.dbo.tbl_item i"
    ELSE
        sqlStr = sqlStr & ",T.outmallprice as sellcash" ''수정. outmallprice
    	sqlStr = sqlStr & " from db_temp.dbo.tbl_xSite_EtcItemLink T"
        sqlStr = sqlStr & " join db_item.dbo.tbl_item i"
    ENd If

    sqlStr = sqlStr & " 	on T.itemid=i.itemid"
    sqlStr = sqlStr & " where mallid='"&sellsite&"'"

    '/제휴상품명과 제휴옵션명만 가지고 매칭하는 제휴몰
    IF GetItemMaeching_itemname_itemoptionname(sellsite) then
        sqlStr = sqlStr & " and ( (outmallitemname='"&Trim(html2db(extitemname))&"' and isnull(outmallitemid,'') = ''))"
        sqlStr = sqlStr & " and T.outmallitemOptionname='"&Replace(extitemoptionname,"'","")&"'"
    ELSE
        sqlStr = sqlStr & " and (outmallitemid='"&xsiteItemID&"') and  (outmallitemid<>'')"

        IF (sellsite="ollehtv") then
            if (extitemoptionname="기본") then
                sqlStr = sqlStr & " and T.outmallitemOptionname=''"
            else
                sqlStr = sqlStr & " and T.outmallitemOptionname='"&Replace(extitemoptionname,"'","")&"'"
            end if
        elseif (sellsite="hanatour") then
            sqlStr = sqlStr & " and T.outmallitemOptionname='"&Replace(extitemoptionname,"'","")&"'"
        elseif (sellsite="gmarket") then
            sqlStr = sqlStr & " and T.outmallitemOptionname='"&Replace(extitemoptionname,"'","")&"'"
        elseif (sellsite="gseshop") then
            sqlStr = sqlStr & " and T.outmallitemOptionname='"&Replace(extitemoptionname,"'","")&"'"
        elseif (sellsite="homeplus") then
            sqlStr = sqlStr & " and T.outmallitemOptionname='"&Replace(extitemoptionname,"'","")&"'"
        elseif (sellsite="ezwel") then
            sqlStr = sqlStr & " and T.outmallitemOptionname='"&Replace(extitemoptionname,"'","")&"'"
        elseif (sellsite="hottracks") then
            sqlStr = sqlStr & " and T.outmallitemOptionname='"&extitemoptionname&"'"
        elseif (sellsite="etsy") then
            sqlStr = sqlStr & " and T.outmallitemOptionname='"&extitemoptionname&"'"
        elseif (sellsite="its29cm") then
            sqlStr = sqlStr & " and T.outmallitemOptionname='"&extitemoptionname&"'"
        elseif (sellsite="itsHiphoper") then
            sqlStr = sqlStr & " and T.outmallitemOptionname='"&extitemoptionname&"'"
        elseif (sellsite="suhaITS") then
            sqlStr = sqlStr & " and T.outmallitemOptionname='"&extitemoptionname&"'"
        elseif (sellsite="ithinksoshop") then
            sqlStr = sqlStr & " and T.outmallitemOptionname='"&extitemoptionname&"'"
        elseif (sellsite="momQ") then
            sqlStr = sqlStr & " and T.outmallitemOptionname='"&extitemoptionname&"'"
        elseif (sellsite="itsMemebox") then
            sqlStr = sqlStr & " and T.outmallitemOptionname='"&extitemoptionname&"'"
        elseif (sellsite="celectory") then
            sqlStr = sqlStr & " and T.outmallitemOptionname='"&extitemoptionname&"'"
        elseif (sellsite="itsKaKaoMakers") then
            sqlStr = sqlStr & " and T.outmallitemOptionname='"&extitemoptionname&"'"
        elseif (sellsite="gsisuper") then
            sqlStr = sqlStr & " and T.outmallitemOptionname='"&extitemoptionname&"'"
        elseif (sellsite="itskakaotalkstore") then
            sqlStr = sqlStr & " and T.outmallitemOptionname='"&extitemoptionname&"'"
        END IF
    END IF

	response.write sqlStr &"<Br>"

    rsget.Open sqlStr,dbget,1
    IF (rsget.RecordCount>1) then
       isTooMany = true

       rw "Too Many Matched ["&xsiteItemID&"]" & extitemname
    else
       if Not rsget.Eof then
            rtitemid = rsget("itemid")
            rtitemoption = rsget("itemoption")
            rtSellPrice = rsget("sellcash")
       end if
    end if
    rsget.close

    if (rtitemid="") then
        rw "No Matched ["&xsiteItemID&"]" & extitemname
    end if
end function

'//제휴몰 상품명을 가지고 텐바이텐 상품명이 있나 확인해서 받아옴
public function getItemIDByUpcheItemCode(sellsite,sellSiteItemID)
    getItemIDByUpcheItemCode = -1
    IF (sellsite="lotteCom") then
        sqlStr = " select top 10 itemid from db_item.dbo.tbl_lotte_regItem"
        sqlStr = sqlStr & " where IsNULL(LotteGoodNo,LotteTmpGoodNo)='"&sellSiteItemID&"'"

        rsget.Open sqlStr,dbget,1

        IF (rsget.RecordCount>1) then
            getItemIDByUpcheItemCode = -1
            rw  sqlStr
        ELSE
            if (Not rsget.EOF) then
        	    getItemIDByUpcheItemCode = rsget("itemid")
        	else
        	    ''19710092-481915 ''중복등록?
        	    IF (sellSiteItemID="19710092") THEN
        	        getItemIDByUpcheItemCode=481915
        	    END IF

        	    rw  sqlStr
        	end if
        ENd IF
        rsget.Close

    ELSEIF (sellsite="interpark") then
        sqlStr = " select top 10 itemid from db_item.dbo.tbl_interpark_reg_Item"
        sqlStr = sqlStr & " where interparkPrdNo='"&sellSiteItemID&"'"

        rsget.Open sqlStr,dbget,1

        IF (rsget.RecordCount>1) then
            getItemIDByUpcheItemCode = -1
        ELSE
            if (Not rsget.EOF) then
        	    getItemIDByUpcheItemCode = rsget("itemid")
        	end if
        ENd IF
        rsget.Close

    ELSEIF (sellsite="cn10x10") then
        sqlStr = " select top 10 itemid from db_item.dbo.tbl_kaffa_reg_item"
        sqlStr = sqlStr & " where itemid='"&sellSiteItemID&"'"

		'response.write sqlStr & "<br>"
        rsget.Open sqlStr,dbget,1

        IF (rsget.RecordCount>1) then
            getItemIDByUpcheItemCode = -1
        ELSE
            if (Not rsget.EOF) then
        	    getItemIDByUpcheItemCode = rsget("itemid")
        	end if
        ENd IF
        rsget.Close
    ENd IF
end function

public function getChrCount(orgStr, delim)
    dim retCNT : retCNT = 0
    dim buf
    buf = split(orgStr,delim)

    if IsArray(buf) then
        retCNT = UBound(buf)
    end if
    getChrCount = retCNT
end function

public function getOptionCodByOptionNameLotte(iitemid,ioptionname)
    dim retStr, sqlStr : retStr=""
    dim IsDoubleOption, IsTreepleOption
    IF (getChrCount(ioptionname,":")=2) THEN
        IF (getChrCount(ioptionname,",")=1) THEN
            IsDoubleOption = TRUE
        END IF
    ELSEIF (getChrCount(ioptionname,":")=3) THEN  '''디자인:c21,폰트선택:폰트2,리필잉크추가 선택:추가안함
        IF (getChrCount(ioptionname,",")=2) THEN
            IsTreepleOption = TRUE
        END IF
    ENd IF

    ioptionname= replace(ioptionname,"'","''")   '' like this CASE : 모델명:SMN-204 you're in
    IF (IsDoubleOption) THEN
        sqlStr = "select top 1 itemoption "
        sqlStr = sqlStr & " from db_item.dbo.tbl_item_option "
        sqlStr = sqlStr & " where itemid="&iitemid&VbcrLF
        sqlStr = sqlStr & " and optionTypename='복합옵션'"
        sqlStr = sqlStr & " and replace(optionname,'*','')='"&SplitValue(SplitValue(ioptionname,",",0),":",1)&","&SplitValue(SplitValue(ioptionname,",",1),":",1)&"'"
    ELSEIF (IsTreepleOption) THEN
        sqlStr = "select top 1 itemoption "
        sqlStr = sqlStr & " from db_item.dbo.tbl_item_option "
        sqlStr = sqlStr & " where itemid="&iitemid&VbcrLF
        sqlStr = sqlStr & " and optionTypename='복합옵션'"
        sqlStr = sqlStr & " and replace(optionname,'*','')='"&SplitValue(SplitValue(ioptionname,",",0),":",1)&","&SplitValue(SplitValue(ioptionname,",",1),":",1)&","&SplitValue(SplitValue(ioptionname,",",2),":",1)&"'"
    ELSE
        sqlStr = "select top 1 itemoption "
        sqlStr = sqlStr & " from db_item.dbo.tbl_item_option "
        sqlStr = sqlStr & " where itemid="&iitemid&VbcrLF
        ''sqlStr = sqlStr & " and optionTypename='"&SplitValue(ioptionname,":",0)&"'"
        sqlStr = sqlStr & " and Replace(Replace(replace(optionname,'*',''),',',''),'#','')=Replace('"&SplitValue(ioptionname,":",1)&"','#','')"
    END IF

	''response.write sqlstr & "<Br>"
	''response.end
    rsget.Open sqlStr,dbget,1
    if (Not rsget.EOF) then
	    retStr = rsget("itemoption")
	end if
    rsget.Close

    If (retStr="") THEN
       ''옵션 매칭이 안되었을때. 수기매칭으로 진행
        sqlStr = "select count(*) as CNT "
        sqlStr = sqlStr & " from db_item.dbo.tbl_item_option "
        sqlStr = sqlStr & " where itemid="&iitemid&VbcrLF
        rsget.Open sqlStr,dbget,1
        if (Not rsget.EOF) then
    	    if (rsget("CNT")>0) THEN retStr = "0000"
    	end if
        rsget.Close

    END IF
    getOptionCodByOptionNameLotte = retStr

if retStr="" then
    rw sqlStr
end if
end function

public function getOptionCodByOptionNameimall(iitemid,ioptionname)
    dim retStr, sqlStr : retStr=""
    dim IsDoubleOption, IsTreepleOption
    IF (getChrCount(ioptionname,":")=2) THEN
        IF (getChrCount(ioptionname,",")=1) THEN
            IsDoubleOption = TRUE
        END IF
    ELSEIF (getChrCount(ioptionname,":")=3) THEN  '''디자인:c21,폰트선택:폰트2,리필잉크추가 선택:추가안함
        IF (getChrCount(ioptionname,",")=2) THEN
            IsTreepleOption = TRUE
        END IF
    ENd IF


    ioptionname= replace(ioptionname,"'","''")   '' like this CASE : 모델명:SMN-204 you're in
    IF (IsDoubleOption) THEN
        sqlStr = "select top 1 itemoption "
        sqlStr = sqlStr & " from db_item.dbo.tbl_item_option "
        sqlStr = sqlStr & " where itemid="&iitemid&VbcrLF
        sqlStr = sqlStr & " and optionTypename='복합옵션'"
        sqlStr = sqlStr & " and replace(optionname,'*','')='"&SplitValue(SplitValue(ioptionname,",",0),":",1)&","&SplitValue(SplitValue(ioptionname,",",1),":",1)&"'"
    ELSEIF (IsTreepleOption) THEN
        sqlStr = "select top 1 itemoption "
        sqlStr = sqlStr & " from db_item.dbo.tbl_item_option "
        sqlStr = sqlStr & " where itemid="&iitemid&VbcrLF
        sqlStr = sqlStr & " and optionTypename='복합옵션'"
        sqlStr = sqlStr & " and replace(optionname,'*','')='"&SplitValue(SplitValue(ioptionname,",",0),":",1)&","&SplitValue(SplitValue(ioptionname,",",1),":",1)&","&SplitValue(SplitValue(ioptionname,",",2),":",1)&"'"
    ELSE
        sqlStr = "select top 1 itemoption "
        sqlStr = sqlStr & " from db_item.dbo.tbl_item_option "
        sqlStr = sqlStr & " where itemid="&iitemid&VbcrLF
        sqlStr = sqlStr & " and optionTypename='"&SplitValue(ioptionname,":",0)&"'"
        sqlStr = sqlStr & " and optionname='"&SplitValue(ioptionname,":",1)&"'"
    END IF
	response.write sqlstr & "<Br>"
'response.end
    rsget.Open sqlStr,dbget,1
    if (Not rsget.EOF) then
	    retStr = rsget("itemoption")
	end if
    rsget.Close

    If (retStr="") THEN
       ''옵션 매칭이 안되었을때. 수기매칭으로 진행
        sqlStr = "select count(*) as CNT "
        sqlStr = sqlStr & " from db_item.dbo.tbl_item_option "
        sqlStr = sqlStr & " where itemid="&iitemid&VbcrLF
        rsget.Open sqlStr,dbget,1
        if (Not rsget.EOF) then
    	    if (rsget("CNT")>=0) THEN retStr = "0000"
    	end if
        rsget.Close

    END IF
    getOptionCodByOptionNameimall = retStr

if retStr="" then
    rw sqlStr
end if
end function

public function getOptionCodByOptionNameGSShop(iitemid,ioptionname)
    dim retStr, sqlStr : retStr=""
    dim IsDoubleOption, IsTreepleOption
    IF (getChrCount(ioptionname,":")=2) THEN
        IF (getChrCount(ioptionname,",")=1) THEN
            IsDoubleOption = TRUE
        END IF
    ELSEIF (getChrCount(ioptionname,":")=3) THEN  '''디자인:c21,폰트선택:폰트2,리필잉크추가 선택:추가안함
        IF (getChrCount(ioptionname,",")=2) THEN
            IsTreepleOption = TRUE
        END IF
    ENd IF


    ioptionname= replace(ioptionname,"'","''")   '' like this CASE : 모델명:SMN-204 you're in
    IF (IsDoubleOption) THEN
        sqlStr = "select top 1 itemoption "
        sqlStr = sqlStr & " from db_item.dbo.tbl_item_option "
        sqlStr = sqlStr & " where itemid='"&iitemid&"' " & VbcrLF
        sqlStr = sqlStr & " and optionTypename='복합옵션'"
        sqlStr = sqlStr & " and replace(optionname,'*','')='"&SplitValue(SplitValue(ioptionname,",",0),":",1)&","&SplitValue(SplitValue(ioptionname,",",1),":",1)&"'"
    ELSEIF (IsTreepleOption) THEN
        sqlStr = "select top 1 itemoption "
        sqlStr = sqlStr & " from db_item.dbo.tbl_item_option "
        sqlStr = sqlStr & " where itemid='"&iitemid&"' " & VbcrLF
        sqlStr = sqlStr & " and optionTypename='복합옵션'"
        sqlStr = sqlStr & " and replace(optionname,'*','')='"&SplitValue(SplitValue(ioptionname,",",0),":",1)&","&SplitValue(SplitValue(ioptionname,",",1),":",1)&","&SplitValue(SplitValue(ioptionname,",",2),":",1)&"'"
    ELSE
        sqlStr = "select top 1 itemoption "
        sqlStr = sqlStr & " from db_item.dbo.tbl_item_option "
        sqlStr = sqlStr & " where itemid='"&iitemid&"' " & VbcrLF
        ''sqlStr = sqlStr & " and optionTypename='"&SplitValue(ioptionname,":",0)&"'"
        sqlStr = sqlStr & " and optionname = '"&ioptionname&"' "
    END IF
'	response.write sqlstr & "<Br>"
'response.end
    rsget.Open sqlStr,dbget,1
    if (Not rsget.EOF) then
	    retStr = rsget("itemoption")
	end if
    rsget.Close

    If (retStr="") THEN
       ''옵션 매칭이 안되었을때. 수기매칭으로 진행
        sqlStr = "select count(*) as CNT "
        sqlStr = sqlStr & " from db_item.dbo.tbl_item_option "
        sqlStr = sqlStr & " where itemid='"&iitemid&"' " & VbcrLF
        rsget.Open sqlStr,dbget,1
        if (Not rsget.EOF) then
    	    if (rsget("CNT")>=0) THEN retStr = "0000"
    	end if
        rsget.Close

    END IF
    getOptionCodByOptionNameGSShop = retStr

if retStr="" then
    rw sqlStr
end if
end function

public function getOptionCodByOptionNameHalfClub(iitemid,ioptionname)
    dim retStr, sqlStr : retStr=""
    dim IsDoubleOption, IsTreepleOption

	If ioptionname = "단일상품" then
		retStr = "0000"
	Else
	    ioptionname= replace(ioptionname,"'","''")   '' like this CASE : 모델명:SMN-204 you're in
	    sqlStr = "select top 1 itemoption "
	    sqlStr = sqlStr & " from db_item.dbo.tbl_item_option "
	    sqlStr = sqlStr & " where itemid='"&iitemid&"' " & VbcrLF
	    sqlStr = sqlStr & " and optionname = '"&ioptionname&"' "
	'	response.write sqlstr & "<Br>"
	'response.end
	    rsget.Open sqlStr,dbget,1
	    if (Not rsget.EOF) then
		    retStr = rsget("itemoption")
		end if
	    rsget.Close

	    If (retStr="") THEN
	       ''옵션 매칭이 안되었을때. 수기매칭으로 진행
	        sqlStr = "select count(*) as CNT "
	        sqlStr = sqlStr & " from db_item.dbo.tbl_item_option "
	        sqlStr = sqlStr & " where itemid='"&iitemid&"' " & VbcrLF
	        rsget.Open sqlStr,dbget,1
	        if (Not rsget.EOF) then
	    	    if (rsget("CNT")>=0) THEN retStr = "0000"
	    	end if
	        rsget.Close

	    END IF
	End IF

 	getOptionCodByOptionNameHalfClub = retStr

	if retStr="" then
	    rw sqlStr
	end if
end function

public function getOptionCodByOptionName11st(iitemid, ioptionname, iMemo)
	Dim retStr, sqlStr : retStr=""
	Dim IsDoubleOption, IsTreepleOption, loops, tmpStr1, tmpStr2, tmpMemo

	tmpStr1 = Split(ioptionname,"-")
	If Ubound(tmpStr1) > 0 Then
		For loops = 0 to Ubound(tmpStr1) - 1
			tmpStr2 = tmpStr2 & tmpStr1(loops) & "-"
		Next
		If Right(tmpStr2,1) = "-" Then
			tmpStr2 = Left(tmpStr2, Len(tmpStr2) - 1)
		End If
		ioptionname = tmpStr2
	End If

	If Instr(ioptionname, "텍스트를 입력하세요:") > 0 Then
		tmpMemo = Trim(Split(ioptionname, "텍스트를 입력하세요:")(1))
	End If

	If getChrCount(tmpMemo, ",") > 0 Then
		iMemo = Split(tmpMemo, ",")(0)
		ioptionname = Split(tmpMemo, iMemo)(1)
	ELSE
		iMemo = tmpMemo
	End If

	If Left(ioptionname, 1) = "," Then
		ioptionname =  Right(ioptionname, Len(ioptionname) - 1)
	End If

	IF (getChrCount(ioptionname,":")=2) THEN
		IF (getChrCount(ioptionname,",")=1) THEN
			IsDoubleOption = TRUE
		END IF
	ELSEIF (getChrCount(ioptionname,":")=3) THEN
		IF (getChrCount(ioptionname,",")=2) THEN
			IsTreepleOption = TRUE
		END IF
	ENd IF

    ioptionname= replace(ioptionname,"'","''")   '' like this CASE : 모델명:SMN-204 you're in
    IF (IsDoubleOption) THEN
        sqlStr = "select top 1 itemoption "
        sqlStr = sqlStr & " from db_item.dbo.tbl_item_option "
        sqlStr = sqlStr & " where itemid='"&iitemid&"' " & VbcrLF
        sqlStr = sqlStr & " and optionTypename='복합옵션'"
        sqlStr = sqlStr & " and replace(optionname,'*','')='"&SplitValue(SplitValue(ioptionname,",",0),":",1)&","&SplitValue(SplitValue(ioptionname,",",1),":",1)&"'"
    ELSEIF (IsTreepleOption) THEN
        sqlStr = "select top 1 itemoption "
        sqlStr = sqlStr & " from db_item.dbo.tbl_item_option "
        sqlStr = sqlStr & " where itemid='"&iitemid&"' " & VbcrLF
        sqlStr = sqlStr & " and optionTypename='복합옵션'"
        sqlStr = sqlStr & " and replace(optionname,'*','')='"&SplitValue(SplitValue(ioptionname,",",0),":",1)&","&SplitValue(SplitValue(ioptionname,",",1),":",1)&","&SplitValue(SplitValue(ioptionname,",",2),":",1)&"'"
    ELSE
        sqlStr = "select top 1 itemoption "
        sqlStr = sqlStr & " from db_item.dbo.tbl_item_option "
        sqlStr = sqlStr & " where itemid='"&iitemid&"' " & VbcrLF
        ''sqlStr = sqlStr & " and optionTypename='"&SplitValue(ioptionname,":",0)&"'"
        sqlStr = sqlStr & " and optionname = '"&SplitValue(ioptionname,":",1)&"' "
    END IF
    rsget.Open sqlStr,dbget,1
    if (Not rsget.EOF) then
	    retStr = rsget("itemoption")
	end if
    rsget.Close

    If (retStr="") THEN
       ''옵션 매칭이 안되었을때. 수기매칭으로 진행
        sqlStr = "select count(*) as CNT "
        sqlStr = sqlStr & " from db_item.dbo.tbl_item_option "
        sqlStr = sqlStr & " where itemid='"&iitemid&"' " & VbcrLF
        rsget.Open sqlStr,dbget,1
        if (Not rsget.EOF) then
    	    if (rsget("CNT")>=0) THEN retStr = "0000"
    	end if
        rsget.Close

    END IF
    getOptionCodByOptionName11st = retStr

if retStr="" then
    rw sqlStr
end if
end function

public function getOptionCodByOptionNameClass(iitemid, ioptionname)
	Dim retStr, sqlStr : retStr=""
	Dim tmpStr1

	tmpStr1 = Split(ioptionname,":")
	If Ubound(tmpStr1) > 0 Then
		ioptionname = Trim(tmpStr1(1))
	End If

	sqlStr = "select top 1 itemoption "
	sqlStr = sqlStr & " from db_item.dbo.tbl_item_option "
	sqlStr = sqlStr & " where itemid='"&iitemid&"' " & VbcrLF
	sqlStr = sqlStr & " and optionname = '"& ioptionname &"' "
    rsget.Open sqlStr,dbget,1
    if (Not rsget.EOF) then
	    retStr = rsget("itemoption")
	end if
    rsget.Close

    If (retStr="") THEN
       ''옵션 매칭이 안되었을때. 수기매칭으로 진행
        sqlStr = "select count(*) as CNT "
        sqlStr = sqlStr & " from db_item.dbo.tbl_item_option "
        sqlStr = sqlStr & " where itemid='"&iitemid&"' " & VbcrLF
        rsget.Open sqlStr,dbget,1
        if (Not rsget.EOF) then
    	    if (rsget("CNT")>=0) THEN retStr = "0000"
    	end if
        rsget.Close

    END IF
    getOptionCodByOptionNameClass = retStr

	if retStr="" then
		rw sqlStr
	end if
end function


Public Function getOptionCodByOptionNameAuction(iitemid,ioptionname, iorderno)
	Dim retStr, sqlStr : retStr=""
	Dim tmpopt, mayOptSuOver
	If ioptionname <> "" Then
		tmpopt = Split(ioptionname,"/")
		If Ubound(tmpopt) = 1 Then
			ioptionname = mid(ioptionname,1,instr(ioptionname,"/")-1)
	        ioptionname = mid(ioptionname,instr(ioptionname,":")+1,100)
			sqlStr = ""
			sqlStr = sqlStr & " SELECT TOP 1 itemoption "
			sqlStr = sqlStr & " FROM db_item.dbo.tbl_item_option "
			sqlStr = sqlStr & " WHERE itemid="&iitemid&VbcrLF
			sqlStr = sqlStr & " and optionname = '"&html2db(ioptionname)&"' "
			'response.write sqlstr & "<Br>"
			'response.end
			rsget.Open sqlStr,dbget,1
			If (Not rsget.EOF) Then
				retStr = rsget("itemoption")
			End If
			rsget.Close
		Else
			retStr = ""
		End If
	End If

	If (retStr = "") Then
		''옵션 매칭이 안되었을때. 수기매칭으로 진행
		sqlStr = ""
		sqlStr = sqlStr & " SELECT count(*) as CNT "
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_item_option "
		sqlStr = sqlStr & " WHERE itemid="&iitemid&VbcrLF
		rsget.Open sqlStr,dbget,1
		If (Not rsget.EOF) Then
			If (rsget("CNT")=0) Then
				retStr = "0000"
			Else
				mayOptSuOver = "Y"
			End If
		End If
		rsget.Close

		If mayOptSuOver = "Y" Then
			sqlStr = ""
			sqlStr = sqlStr & " SELECT TOP 1 replace(orderItemOption, 'FF', '') as orderItemOption "
			sqlStr = sqlStr & " FROM db_temp.dbo.tbl_xSite_TmpOrder "
			sqlStr = sqlStr & " WHERE OutMallOrderSerial = '"&Trim(iorderno)&"' "
			sqlStr = sqlStr & " and Left(orderItemOption, 2) = 'FF' "
			sqlStr = sqlStr & " ORDER BY orderItemOption DESC "
			rsget.Open sqlStr,dbget,1
			If (Not rsget.EOF) Then
				retStr = "FF" & CInt(rsget("orderItemOption")) + 1
			Else
				retStr = "FF10"
			End If
			rsget.Close
		End If
	END IF

'rw retStr
'response.end

	getOptionCodByOptionNameAuction = retStr

	If retStr="" Then
		rw sqlStr
	End If
End Function

'//제휴몰 옵션명을 가지고 텐바이텐 옵션이 있나 확인해서 받아옴
public function getOptionCodByOption(iitemid,ioption)
    dim retStr, sqlStr : retStr=""

    sqlStr = "select top 1 itemoption "
    sqlStr = sqlStr & " from db_item.dbo.tbl_item_option "
    sqlStr = sqlStr & " where itemid='"&iitemid&"'"
    sqlStr = sqlStr & " and itemoption='"&ioption&"'"

	'response.write sqlstr & "<Br>"
    rsget.Open sqlStr,dbget,1
    if (Not rsget.EOF) then
	    retStr = rsget("itemoption")
	end if
    rsget.Close

    getOptionCodByOption = retStr
end function

public function getOptionCodeByMakeShopOptCode(iitemid,imakeshopoptcode)
    dim retStr, sqlStr : retStr=""
    sqlStr = "select top 1 tenoptioncode "
    sqlStr = sqlStr & " from db_item.[dbo].[tbl_makeglob_product_option] "
    sqlStr = sqlStr & " where product_code='"&iitemid&"'"&VbcrLF
    sqlStr = sqlStr & " and idx='"&imakeshopoptcode&"'"

	response.write sqlstr & "<Br>"
    rsget.Open sqlStr,dbget,1
    if (Not rsget.EOF) then
	    retStr = rsget("tenoptioncode")
	end if
    rsget.Close

    If (retStr="") THEN
       ''옵션 매칭이 안되었을때. 수기매칭으로 진행
        sqlStr = "select count(*) as CNT "
        sqlStr = sqlStr & " from db_item.dbo.tbl_item_option "
        sqlStr = sqlStr & " where itemid='"&iitemid&"' " & VbcrLF
        rsget.Open sqlStr,dbget,1
        if (Not rsget.EOF) then
    	    if (rsget("CNT")>=0) THEN retStr = "0000"
    	end if
        rsget.Close

    END IF
	getOptionCodeByMakeShopOptCode = retStr
End Function

'11번가 옵션 코드 얻기
Public Function get11stOptionCodeByOptionName(iitemid, ioptionname)
	Dim retStr, sqlStr : retStr=""
    sqlStr = "SELECT TOP 1 itemoption "
    sqlStr = sqlStr & " FROM db_item.[dbo].[tbl_item_multiLang_option] "
    sqlStr = sqlStr & " WHERE itemid = '"&iitemid&"'"
    sqlStr = sqlStr & " and optionname='"&TRIM(html2db(ioptionname))&"'"
    sqlStr = sqlStr & " and countryCd = 'EN' "
    rsget.Open sqlStr,dbget,1
    If (Not rsget.EOF) Then
	    retStr = rsget("itemoption")
	End If
    rsget.Close

	get11stOptionCodeByOptionName = retStr
End Function

'질링고 상품, 옵션 코드 얻기
Public Function getItemidOptionCodeByZilignoGoodno(izilingoGoodno)
	Dim retStr, sqlStr
	sqlStr = ""
    sqlStr = sqlStr & " SELECT TOP 1 itemid, itemoption "
    sqlStr = sqlStr & " FROM db_etcmall.dbo.tbl_zilingo_regItem "
    sqlStr = sqlStr & " WHERE zilingoGoodNo = '"& izilingoGoodno &"' "
    rsget.Open sqlStr,dbget,1
    If (Not rsget.EOF) Then
	    retStr = rsget("itemid") & "||" & rsget("itemoption")
	End If
    rsget.Close
	getItemidOptionCodeByZilignoGoodno = retStr
End Function

public function getOptionCodByOptionName(iitemid,ioptionname)
    dim retStr, sqlStr : retStr=""
    sqlStr = "select top 1 itemoption "
    sqlStr = sqlStr & " from db_item.dbo.tbl_item_option "
    sqlStr = sqlStr & " where itemid='"&iitemid&"'"&VbcrLF
    sqlStr = sqlStr & " and optionname='"&TRIM(html2db(ioptionname))&"'"

	'response.write sqlstr & "<Br>"
    rsget.Open sqlStr,dbget,1
    if (Not rsget.EOF) then
	    retStr = rsget("itemoption")
	end if
    rsget.Close

    if (retStr="") then
        sqlStr = "select top 1 itemoption "
        sqlStr = sqlStr & " from db_item.dbo.tbl_item_option "
        sqlStr = sqlStr & " where itemid="&iitemid&VbcrLF
        sqlStr = sqlStr & " and replace(replace(optionname,',',''),' ','')='"&TRIM(replace(replace(html2db(ioptionname),",","")," ",""))&"'"
        rsget.Open sqlStr,dbget,1
        if (Not rsget.EOF) then
    	    retStr = rsget("itemoption")
    	end if
        rsget.Close
    end if
    getOptionCodByOptionName = retStr

if retStr="" then
    rw sqlStr
end if
'    if (retStr="") then
'        sqlStr = "select top 1 itemoption "
'        sqlStr = sqlStr & " from db_item.dbo.tbl_item_option "
'        sqlStr = sqlStr & " where itemid="&iitemid&VbcrLF
'        sqlStr = sqlStr & " and optionname like '"&ioptionname&"%'"
'    end if
end function

public function getOptionCodByOptionNameSSG(iitemid,ioptionname)
    dim retStr, sqlStr : retStr=""
    sqlStr = "select top 1 itemoption "
    sqlStr = sqlStr & " from db_item.dbo.tbl_item_option "
    sqlStr = sqlStr & " where itemid='"&iitemid&"'"&VbcrLF
    sqlStr = sqlStr & " and optionname='"&TRIM(html2db(ioptionname))&"'"
	'response.write sqlstr & "<Br>"
	rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
    if (Not rsget.EOF) then
	    retStr = rsget("itemoption")
	end if
    rsget.Close

    If (retStr="") THEN
       ''옵션 매칭이 안되었을때. 수기매칭으로 진행
        sqlStr = "select count(*) as CNT "
        sqlStr = sqlStr & " from db_item.dbo.tbl_item_option "
        sqlStr = sqlStr & " where itemid='"&iitemid&"' " & VbcrLF
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
        if (Not rsget.EOF) then
    	    if (rsget("CNT")>=0) THEN retStr = "0000"
    	end if
        rsget.Close
    end if
    getOptionCodByOptionNameSSG = retStr
end function

Class CxSiteTempLinkSubItem
    public Fitemid
    public Fitemoption
    public FmallID
    public Foutmallitemid
    public Foutmallitemname
    public FoutmallitemOptionname
    public FoutmallPrice
    public FoutmallSellYn
    public Fitemname
    public FitemOptionName
    public Fsellyn
    public Fsellcash
    public FoptAddPrice
    public Fsmallimage

    public Flimityn
    public Flimitno
    public Flimitsold

    public Foptusing
    public Foptsellyn
    public Foptlimityn
    public Foptlimitno
    public Foptlimitsold

    public function IsOptionSoldout
        IsOptionSoldout = false
        if (Fitemoption="0000") then Exit function

        IsOptionSoldout = (Foptusing="N") or (Foptsellyn<>"Y") or ((Foptlimityn="Y") and (Foptlimitno-Foptlimitsold<1))

    end function

    public function IsLimitSell
        IsLimitSell = (Flimityn="Y")
    end function

    public function getLimitRemainNo()
        dim ret
        if (Fitemoption="0000") then
            ret = Flimitno-Flimitsold
        else
            ret = Foptlimitno-Foptlimitsold
        end if

        if ret<1 then ret=0
        getLimitRemainNo = ret
    end function

    Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
End Class

Class CxSiteTempLinkItem
    public FItemList()
	public FOneItem
	public FCurrPage
	public FPageSize
	public FResultCount
	public FScrollCount
	public FTotalCount
	public FTotalPage
	public FRectSellSite
	public FRectItemid
	public FRectItemOption
	public FRectStateDiff
	public FRectPriceDiff
    public FRectitemidarr
    public FRectoutmallitemidarr

    public function getOnexSiteTempLinkItem()
        dim sqlStr
        sqlStr = "select T.itemid, T.itemoption, T.mallID, T.outmallitemid"
        sqlStr = sqlStr& " ,T.outmallitemname, T.outmallPrice, T.outmallSellYn , i.itemname,i.sellyn,i.sellcash"
        sqlStr = sqlStr& " ,T.outmallitemOptionname"
        sqlStr = sqlStr& " ,isNULL(o.optionname,'') as itemoptionname"
        sqlStr = sqlStr& " from db_temp.dbo.tbl_xSite_EtcItemLink T"
        sqlStr = sqlStr& "      left join db_item.dbo.tbl_item i"
        sqlStr = sqlStr& "      on T.itemid=i.itemid"
        sqlStr = sqlStr& "      left join db_item.dbo.tbl_item_option o"
        sqlStr = sqlStr& "      on T.itemid=o.itemid"
        sqlStr = sqlStr& "      and T.itemoption=o.itemoption"
        sqlStr = sqlStr& " where 1=1"
        sqlStr = sqlStr& " and T.itemid="&FRectItemid&""
        sqlStr = sqlStr& " and T.itemoption='"&FRectItemOption&"'"
        sqlStr = sqlStr& " and T.mallID='"&FRectSellSite&"'"
        sqlStr = sqlStr& " order by T.itemid desc"

		'response.write sqlstr & "<Br>"
        rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

		FResultCount = rsget.RecordCount

		if  not rsget.EOF  then
			set FOneItem = new CxSiteTempLinkSubItem

            FOneItem.Fitemid          = rsget("itemid")
            FOneItem.Fitemoption      = rsget("itemoption")
            FOneItem.FmallID          = rsget("mallID")
            FOneItem.Foutmallitemid   = rsget("outmallitemid")
            FOneItem.Foutmallitemname = rsget("outmallitemname")
            FOneItem.FoutmallitemOptionname = rsget("outmallitemOptionname")
            FOneItem.FoutmallPrice    = rsget("outmallPrice")
            FOneItem.FoutmallSellYn   = rsget("outmallSellYn")
            FOneItem.Fitemname        = rsget("itemname")
            FOneItem.Fitemoptionname  = rsget("itemoptionname")
            FOneItem.Fsellyn          = rsget("sellyn")
            FOneItem.Fsellcash        = rsget("sellcash")
		end if
		rsget.Close
    end function

    public function xSiteTempLinkItemList()
        dim sqlStr
        dim addSQL

        addSQL = ""

        if (FRectSellSite<>"") then
            addSQL = addSQL & " and T.mallID='"&FRectSellSite&"'"
        end if

        if (FRectStateDiff<>"") then
            addSQL = addSQL & " and (( (T.outmallSellYn='N' OR T.outmallSellYn='X') and isNULL(i.sellyn,'Y') in ('Y'))"
            addSQL = addSQL & "     or (T.outmallSellYn='Y' and (isNULL(i.limityn,'N')='Y') and (isNULL(i.limitno,0)-isNULL(i.limitsold,0)<1))"
            addSQL = addSQL & "     or (T.outmallSellYn='Y' and isNULL(i.sellyn,'Y') in ('S','N'))"
            addSQL = addSQL & "     or (T.outmallSellYn='Y' and (isNULL(o.isusing,'Y')='N' or isNULL(o.optsellyn,'Y')='N'))"
            addSQL = addSQL & "     or (T.outmallSellYn='Y' and (isNULL(o.optlimityn,'N')='Y' and (isNULL(o.optlimitno,0)-isNULL(o.optlimitsold,0)<1)))"
            addSQL = addSQL & " )"
        end if

        if (FRectPriceDiff<>"") then
            addSQL = addSQL & " and (T.outmallPrice<>isNULL(i.sellcash,0)+isNULL(o.optAddPrice,0))"
        end if

        if (FRectitemidarr<>"") then
            FRectitemidarr = Trim(FRectitemidarr)
            if Right(FRectitemidarr,1)="," then FRectitemidarr=Left(FRectitemidarr,Len(FRectitemidarr)-1)
            addSQL = addSQL & " and T.itemid in ("&FRectitemidarr&")"
        end if

        if (FRectoutmallitemidarr<>"") then
            FRectoutmallitemidarr = Trim(FRectoutmallitemidarr)
            if Right(FRectoutmallitemidarr,1)="," then FRectoutmallitemidarr=Left(FRectoutmallitemidarr,Len(FRectoutmallitemidarr)-1)
            FRectoutmallitemidarr = replace(FRectoutmallitemidarr,",","','")
            addSQL = addSQL & " and T.outmallitemid in ('"&FRectoutmallitemidarr&"')"
        end if

        sqlStr = "select count(*) as CNT , CEILING(CAST(Count(*) AS FLOAT)/" & FPageSize & ") as totPg"
        sqlStr = sqlStr& " from db_temp.dbo.tbl_xSite_EtcItemLink T"
        sqlStr = sqlStr& "      left join db_item.dbo.tbl_item i"
        sqlStr = sqlStr& "      on T.itemid=i.itemid"
        sqlStr = sqlStr& "      left join db_item.dbo.tbl_item_option o"
        sqlStr = sqlStr& "      on T.itemid=o.itemid"
        sqlStr = sqlStr& "      and T.itemoption=o.itemoption"
        sqlStr = sqlStr& " where 1=1"
        sqlStr = sqlStr& addSQL

		'response.write sqlstr & "<Br>"
        rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
		rsget.Close

		'지정페이지가 전체 페이지보다 클 때 함수종료
		if Cint(FCurrPage)>Cint(FTotalPage) then
			FResultCount = 0
			exit function
		end if

        sqlStr = "select top "&(FCurrPage*FPageSize)&" T.itemid, T.itemoption, T.mallID, T.outmallitemid"
        sqlStr = sqlStr& " ,T.outmallitemname, T.outmallPrice, T.outmallSellYn , i.itemname,i.sellyn,i.sellcash"
        sqlStr = sqlStr& " ,T.outmallitemOptionname"
        sqlStr = sqlStr& " ,isNULL(o.optionname,'') as itemoptionname"
        sqlStr = sqlStr& " ,isNULL(o.optAddPrice,0) as optAddPrice"
        sqlStr = sqlStr& " ,i.smallimage,i.limityn,i.limitno,i.limitsold"
        sqlStr = sqlStr& " ,o.isusing as optusing,o.optsellyn,o.optlimityn,o.optlimitno,o.optlimitsold"
        sqlStr = sqlStr& " from db_temp.dbo.tbl_xSite_EtcItemLink T"
        sqlStr = sqlStr& "      left join db_item.dbo.tbl_item i"
        sqlStr = sqlStr& "      on T.itemid=i.itemid"
        sqlStr = sqlStr& "      left join db_item.dbo.tbl_item_option o"
        sqlStr = sqlStr& "      on T.itemid=o.itemid"
        sqlStr = sqlStr& "      and T.itemoption=o.itemoption"
        sqlStr = sqlStr& " where 1=1"
        sqlStr = sqlStr& addSQL
        sqlStr = sqlStr& " order by T.itemid desc, T.itemoption"

		'response.write sqlstr & "<Br>"
        rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CxSiteTempLinkSubItem

                FItemList(i).Fitemid          = rsget("itemid")
                FItemList(i).Fitemoption      = rsget("itemoption")
                FItemList(i).FmallID          = rsget("mallID")
                FItemList(i).Foutmallitemid   = rsget("outmallitemid")
                FItemList(i).Foutmallitemname = rsget("outmallitemname")
                FItemList(i).FoutmallitemOptionname = rsget("outmallitemOptionname")
                FItemList(i).FoutmallPrice    = rsget("outmallPrice")
                FItemList(i).FoutmallSellYn   = rsget("outmallSellYn")
                FItemList(i).Fitemname        = rsget("itemname")
                FItemList(i).Fitemoptionname  = rsget("itemoptionname")
                FItemList(i).Fsellyn          = rsget("sellyn")
                FItemList(i).Fsellcash        = rsget("sellcash")

                FItemList(i).FoptAddPrice     = rsget("optAddPrice")
                FItemList(i).Fsmallimage      = rsget("smallimage")

                if Not(FItemList(i).FsmallImage="" or isNull(FItemList(i).FsmallImage)) then
					FItemList(i).FsmallImage = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + FItemList(i).FsmallImage
				else
					FItemList(i).FsmallImage = "http://fiximage.10x10.co.kr/images/spacer.gif"
				end if

                FItemList(i).Flimityn       = rsget("limityn")
                FItemList(i).Flimitno       = rsget("limitno")
                FItemList(i).Flimitsold     = rsget("limitsold")

				FItemList(i).Foptusing      = rsget("optusing")
                FItemList(i).Foptsellyn     = rsget("optsellyn")
                FItemList(i).Foptlimityn    = rsget("optlimityn")
                FItemList(i).Foptlimitno    = rsget("optlimitno")
                FItemList(i).Foptlimitsold  = rsget("optlimitsold")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close

    end function

    Private Sub Class_Initialize()
		redim  FItemList(0)

		FCurrPage =1
		FPageSize = 12
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
		FTotalPage =0
	End Sub

	Private Sub Class_Terminate()

	End Sub

	public Function HasPreScroll()
		HasPreScroll = StartScrollPage > 1
	end Function

	public Function HasNextScroll()
		HasNextScroll = FTotalPage > StartScrollPage + FScrollCount -1
	end Function

	public Function StartScrollPage()
		StartScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
	end Function
end Class

Class CxSiteTempOrderItem
	public forderemail
	public FOutMallOrderSeq
''    public Fcompanyid
    public FSellSite
    public FSellSiteName
    public FmatchItemID
    public FmatchItemOption
    public FmatchItemName
    public FmatchItemOptionName
    public ForderItemID
    public ForderItemName
    public ForderItemOption
    public ForderItemOptionName
    public FOutMallOrderSerial
    public FfoundPrdcode
    public Ftplprdcode
    public Ftplorderserial
    public FmatchState
    public FOrderName
    public FOrderTelNo
    public FOrderHpNo
    public FReceiveName
    public FReceiveTelNo
    public FReceiveHpNo
    public FReceiveZipCode
    public FReceiveAddr1
    public FReceiveAddr2
    public Fdeliverymemo
    public FSellPrice
    public FRealSellPrice
    public FItemOrderCount
    public FrequireDetail
    public ForderDlvPay
    public FSelldate
    public Forderserial
    public FoptionCnt
    public FDuppExists
    public FaddDlvExists
    public FOrgDetailKey
    public FordercsGbn
    public FRef_OutMallOrderSerial
	public fcountryCode
    public Fsellyn
    public Flimityn
    public Flimitno
    public Flimitsold
    public FSellcash
    public Foptusing
    public Foptsellyn
    public Foptlimityn
    public Foptlimitno
    public Foptlimitsold
    public Foptaddprice
	public fitemid
	public Fitemoption
	public fmallID
	public foutmallitemid
	public foutmallitemname
	public foutmallitemOptionname
	public foutmallPrice
	public foutmallSellYn
	public fitemname
	public fsmallimage
	public fsoldout
    public Fbeadaldiv
	public FItemdiv
	public FPaydate
	public FFFExists
	public FShopifyOrderName
	public FOverseasPrice

    public function isCurDiffPrice()
        isCurDiffPrice = FSellcash + Foptaddprice<>FSellPrice
    end function

    public function getCurDiffPriceHtml
        getCurDiffPriceHtml = FormatNumber(FSellcash,0)
        if (Foptaddprice<>0) then
            getCurDiffPriceHtml = getCurDiffPriceHtml + " + " + FormatNumber(Foptaddprice,0)
        end if
    end function

    public function isCurItemSoldOut()
        isCurItemSoldOut = ((Fsellyn<>"Y") or ((Flimityn="Y") and (Flimitno-Flimitsold<1)))
    end function

    public function isCurItemOptionSoldOut()
        isCurItemOptionSoldOut = ((Foptusing<>"Y") or (Foptsellyn<>"Y") or ((Foptlimityn="Y") and (Foptlimitno-Foptlimitsold<1)))
    end function

    public function isCancelOrder()
        isCancelOrder = (FordercsGbn="2")
    end function

    public function getOrderCsGbnName()
        IF IsNULL(FordercsGbn) then
            getOrderCsGbnName=""
            exit function
        end if

        IF CStr(FordercsGbn)="0" then
            getOrderCsGbnName=""
            exit function
        end if

        if FordercsGbn="3" then
            getOrderCsGbnName="<font color=blue>CS</font>"
        elseif FordercsGbn="8" then                             ''주문입력시 기입력 인경우 이 플래그 업데이트
            getOrderCsGbnName="<font color=blue>중복</font>"
        elseif FordercsGbn="2" then  ''cjMall 주문취소
            getOrderCsGbnName="<font color=red>취소</font>"
        else
            getOrderCsGbnName=CStr(FordercsGbn)
        end if
    end function

	public function getmatchStateString()
		if FmatchState="I" then
			'단품
			getmatchStateString = "엑셀입력"
		elseif FmatchState="P" then
			'제외
			getmatchStateString = "상품매칭완료"
		elseif FmatchState="O" then
			'포함
			getmatchStateString = "주문입력완료"
		end if
	end function

    public function getNotiStateString()
        if (FmatchState="0") then
            getNotiStateString = "<font color=blue>수신</font>"
        elseif (FmatchState="3") then
            getNotiStateString = "처리"
        elseif (FmatchState="9") then
            getNotiStateString = "입력전취소"
        else
            getNotiStateString = FmatchState
        end if
    end function

	public function getorderItemName()
		''if (Fcompanyid="toms" and CStr(FSellSite) = "5") then
		''	'탐스 - 후이즈쇼핑몰
		''	getorderItemName = Left(ForderItemName, (Len(ForderItemName) - 3))
		''else
			getorderItemName = ForderItemName
		''end if
	end function

    public function IsItemOptionNameNotMatched()
        dim buf
        if (FSellSite="lotteCom") or  (FSellSite="lotteimall") then
            if InStr(ForderItemOptionName,":")>0 then
                buf = Mid(ForderItemOptionName,InStr(ForderItemOptionName,":")+1,255)
                if InStr(buf,":")>0 then ''이중옵션
                    if InStr(buf,",")>0 then
                        IsItemOptionNameNotMatched = Trim(left(buf,InStr(buf,","))+Mid(buf,InStr(buf,":")+1,255))<>Trim(FmatchItemOptionName)
                    else
                        IsItemOptionNameNotMatched = false
                    end if
                else
                    IsItemOptionNameNotMatched = Trim(buf)<>Trim(FmatchItemOptionName)
                end if
            else
                IsItemOptionNameNotMatched = Trim(ForderItemOptionName)<>Trim(FmatchItemOptionName)
            end if
        elseif (FSellSite="interpark") then
            if InStr(ForderItemOptionName,"/")>0 then
                buf = Mid(ForderItemOptionName,InStr(ForderItemOptionName,"/")+1,255)
                if InStr(buf,"/")>0 then ''이중옵션
                    if InStr(buf,",")>0 then
                        IsItemOptionNameNotMatched = Trim(left(buf,InStr(buf,","))+Mid(buf,InStr(buf,"/")+1,255))<>Trim(FmatchItemOptionName)
                    else
                        IsItemOptionNameNotMatched = false
                    end if
                else
                    IsItemOptionNameNotMatched = Trim(buf)<>Trim(FmatchItemOptionName)
                end if
            else
                IsItemOptionNameNotMatched = Trim(ForderItemOptionName)<>Trim(FmatchItemOptionName)
            end if
        elseif (FSellSite="lotteimall") then
            IsItemOptionNameNotMatched = Trim(ForderItemOptionName)<>Replace(Trim(FmatchItemOptionName),",","")
        end if

        if (IsItemOptionNameNotMatched) then
            if (FSellSite="lotteimall") then
                IsItemOptionNameNotMatched = Trim(ForderItemOptionName)<>Replace(Trim(FmatchItemOptionName),",","/")
            end if
        end if
    end function

    public function IsItemOptionNotMatched()

        if isNULL(Fmatchitemoption) then
            IsItemOptionNotMatched = true
	        exit function
	    end if

        if isNULL(FmatchItemOptionName) then
            IsItemOptionNotMatched = true
	        exit function
	    end if

	    if (Fmatchitemoption<>"0000") and (FmatchItemOptionName="") then
	        IsItemOptionNotMatched = true
	        exit function
	    end if

'	    if (Fmatchitemoption<>"0000") and (Fmatchitemoption<>ForderItemOption) then
'	        IsItemOptionNotMatched = true
'	        exit function
'	    end if

	    IsItemOptionNotMatched = false

    end function

	public function IsItemMatched()
	    if isNULL(Fmatchitemid) or isNULL(Fmatchitemoption) then
	        IsItemMatched = false
	        exit function
	    end if

	    if IsNULL(FmatchItemName) or isNULL(FmatchItemOptionName) then
	        IsItemMatched = false
	        exit function
	    end if

		if (FmatchState="I") then
			IsItemMatched = true
		else
			IsItemMatched = false
		end if

	end function

    public function IsCjMallStarCASE()
    	exit function
        IsCjMallStarCASE = ((FSellSite="cjmall") and ((InStr(FOrderName,"*")>0) or (InStr(fReceiveName,"*")>0)))
    end function

    Private Sub Class_Initialize()
	End Sub

	Private Sub Class_Terminate()
	End Sub
end Class

Class CxSiteTempOrder
    public FItemList()
	public FOneItem
	public FCurrPage
	public FPageSize
	public FResultCount
	public FScrollCount
	public FTotalCount
	public FTotalPage
	public FRectCompanyID
	public FRectSellSite
	public FRectMatchState
	public FRectCsViewYn
	public FRectOverseaViewYn
	public FRectorderserial
	public FRectOutMallOrderSerial
	public FRectOutMallOrderSeq
	public FEMSPrice

	public frectitemid
	public FrectItemName
	public frectoutmallSellYn
	public FrectSoldOUT
	public frectmallid
	public FRectregYYYYMMDD
	public FRectInc3pl

	'/admin/etc/orderinput/xSiteOrderedit.asp
	public sub fxsiteorderedit()
		dim sqlStr, sqlsearch
		if frectoutmallorderseq="" then exit sub

		if frectoutmallorderseq<>"" then
			sqlsearch = sqlsearch & " and outmallorderseq='"& frectoutmallorderseq &"'" + vbcrlf
		end if

		sqlStr = "select top " & Cstr(FPageSize * FCurrPage) + vbcrlf
		sqlStr = sqlStr & " sellsite, isNULL(OrderSerial,'') as OrderSerial, OutMallOrderSerial, OrderName, ReceiveName" + vbcrlf
		sqlStr = sqlStr & " from db_temp.dbo.tbl_xSite_TMPOrder" + vbcrlf
		sqlStr = sqlStr & " where 1=1 " & sqlsearch + vbcrlf

		'response.write sqlStr &"<br>"
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

		ftotalcount = rsget.recordcount
		FResultCount = rsget.recordcount

		if  not rsget.EOF  then
			set foneitem = new CxSiteTempOrderItem

			foneitem.fsellsite = rsget("sellsite")
			foneitem.fOrderSerial = rsget("OrderSerial")
			foneitem.fOutMallOrderSerial = rsget("OutMallOrderSerial")
			foneitem.fOrderName = db2html(rsget("OrderName"))
			foneitem.fReceiveName = db2html(rsget("ReceiveName"))
		end if
		rsget.Close
	end sub

	'//admin/etc/orderinput/soldout_scheduler.asp
	public sub getxsitesoldout_scheduler
		Dim sqlStr, i, sqlsearch

		If frectitemid <> "" Then
			sqlsearch = sqlsearch & " AND l.itemid = '" & frectitemid & "' "
		End If

		If FrectItemName <> "" Then
			sqlsearch = sqlsearch & " AND I.itemname like '%" & FrectItemName & "%' "
		End If

		If frectoutmallSellYn <> "" Then
			sqlsearch = sqlsearch & " AND l.outmallSellYn = '" & frectoutmallSellYn & "' "
		End If

		If FrectSoldOUT = "Y" Then
			sqlsearch = sqlsearch & " AND I.sellyn <> 'Y' "
		ElseIf FrectSoldOUT = "N" Then
			sqlsearch = sqlsearch & " AND I.sellyn = 'Y' "
		End If

		if frectmallid <> "" then
			sqlsearch = sqlsearch & " and l.mallid ='"&frectmallid&"'"
		end if

		sqlStr = "select count(*) as cnt"
		sqlStr = sqlStr & " from db_temp.dbo.tbl_xSite_EtcItemLink l"
		sqlStr = sqlStr & " join [db_item].[dbo].[tbl_item] AS I"
		sqlStr = sqlStr & " 	ON l.itemid = I.itemid"
		sqlStr = sqlStr & " where 1=1 " & sqlsearch

		'response.write sqlStr & "<br>"
		rsget.Open sqlStr, dbget ,1
			ftotalcount = rsget("cnt")
		rsget.Close

		sqlStr = "select top " & (FPageSize * FCurrPage)
		sqlStr = sqlStr & " l.itemid, l.itemoption, l.mallID, l.outmallitemid, l.outmallitemname, l.outmallitemOptionname"
		sqlStr = sqlStr & " , l.outmallPrice, l.outmallSellYn"
		sqlStr = sqlStr & " , I.smallimage, I.itemname, I.limitno, I.limitsold, I.sellyn, I.limityn"
		sqlStr = sqlStr & " from db_temp.dbo.tbl_xSite_EtcItemLink l"
		sqlStr = sqlStr & " join [db_item].[dbo].[tbl_item] AS I"
		sqlStr = sqlStr & " 	ON l.itemid = I.itemid"
		sqlStr = sqlStr & " where 1=1 " & sqlsearch
		sqlStr = sqlStr & " order by l.mallID asc, l.itemid desc"

		'response.write sqlStr & "<br>"
		rsget.Open sqlStr, dbget ,1

		if (FCurrPage * FPageSize < FTotalCount) then
			FResultCount = FPageSize
		else
			FResultCount = FTotalCount - FPageSize*(FCurrPage-1)
		end if

		FTotalPage = (FTotalCount\FPageSize)

		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

		redim preserve FItemList(FResultCount)

		rsget.PageSize= FPageSize
		If  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			Do Until rsget.Eof
				set FItemList(i) = new CxSiteTempOrderItem

				FItemList(i).fitemid		= rsget("itemid")
				FItemList(i).fitemoption		= rsget("itemoption")
				FItemList(i).fmallID		= rsget("mallID")
				FItemList(i).foutmallitemid		= rsget("outmallitemid")
				FItemList(i).foutmallitemname		= rsget("outmallitemname")
				FItemList(i).foutmallitemOptionname		= rsget("outmallitemOptionname")
				FItemList(i).foutmallPrice		= rsget("outmallPrice")
				FItemList(i).foutmallSellYn		= rsget("outmallSellYn")
				FItemList(i).fitemname		= rsget("itemname")
				FItemList(i).fsmallimage	= webImgUrl & "/image/small/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + rsget("smallimage")
				IF rsget("limitno")<>"" and rsget("limitsold")<>"" Then
					FItemList(i).fsoldout = (rsget("sellyn")<>"Y") or ((rsget("limityn") = "Y") and (clng(rsget("limitno"))-clng(rsget("limitsold"))<1))
				Else
					FItemList(i).fsoldout = (rsget("sellyn")<>"Y")
				End If
				If (rsget("sellyn") = "S") Then
					FItemList(i).fsoldout = (rsget("sellyn") = "S")
				End IF

				i=i+1
				rsget.moveNext
			Loop
		End If
		rsget.Close
	end sub

	public function getOrderNotiLogList()
	    dim i,sqlStr
	    sqlStr = "select count(*) as CNT , CEILING(CAST(Count(*) AS FLOAT)/" & FPageSize & ") as totPg"
	    sqlStr = sqlStr & " from db_temp.dbo.tbl_LTiMall_OrdNoti t"
	    sqlStr = sqlStr & " where 1=1"
	    if (FRectOutMallOrderSerial<>"") then
    	    sqlStr = sqlStr & " and ORDER_NO='"&FRectOutMallOrderSerial&"'"
    	end if

		'response.write sqlstr & "<Br>"
    	rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
		rsget.Close

		'지정페이지가 전체 페이지보다 클 때 함수종료
		if Cint(FCurrPage)>Cint(FTotalPage) then
			FResultCount = 0
			exit function
		end if

		sqlStr = "select top " + CStr(FPageSize*FCurrPage) + " t.* "
	    sqlStr = sqlStr & " from db_temp.dbo.tbl_LTiMall_OrdNoti t"
	    sqlStr = sqlStr & " where 1=1"
	    if (FRectOutMallOrderSerial<>"") then
    	    sqlStr = sqlStr & " and ORDER_NO='"&FRectOutMallOrderSerial&"'"
    	end if
    	sqlStr = sqlStr & " order by t.notistatus, t.regdate desc"
    	''sqlStr = sqlStr & " order by t.ORDER_NO desc, t.ORDER_SEQ"

		'response.write sqlStr & "<Br>"
	    rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CxSiteTempOrderItem

    			''FItemList(i).FOutMallOrderSeq		= rsget("OutMallOrderSeq")
    			FItemList(i).FSellSite				= "lotteimall"
    			FItemList(i).FSellSiteName	        = "롯데iMall"
                FItemList(i).FmatchItemID           = ""
                FItemList(i).FmatchItemOption       = ""
    			FItemList(i).FmatchItemName			= ""
    			FItemList(i).FmatchItemOptionName	= ""
                FItemList(i).ForderItemID			= splitValue(rsget("ENTP_DT_CODE"),"_",0)
    			FItemList(i).ForderItemName			= rsget("GOODS_NAME")
    			FItemList(i).ForderItemOption		= splitValue(rsget("ENTP_DT_CODE"),"_",1)
    			FItemList(i).ForderItemOptionName	= rsget("GOODSDT_INFO")
    			FItemList(i).FOutMallOrderSerial	= rsget("ORDER_NO")
    			''FItemList(i).FfoundPrdcode			= rsget("foundPrdcode")
    			''FItemList(i).Ftplorderserial		= rsget("tplorderserial")
    			FItemList(i).FmatchState			= rsget("notistatus")
                FItemList(i).FOrderName             = db2HTML(rsget("O_NAME"))
                FItemList(i).FOrderTelNo            = db2HTML(rsget("O_TEL"))
                FItemList(i).FOrderHpNo             = db2HTML(rsget("O_HTEL"))
                FItemList(i).FReceiveName           = db2HTML(rsget("S_NAME"))
                FItemList(i).FReceiveTelNo          = db2HTML(rsget("S_TEL"))
                FItemList(i).FReceiveHpNo           = db2HTML(rsget("S_HTEL"))
                FItemList(i).FReceiveZipCode        = db2HTML(rsget("S_POST"))
                FItemList(i).FReceiveAddr1          = db2HTML(rsget("S_ADDR"))
                FItemList(i).FReceiveAddr2          = ""
                FItemList(i).Fdeliverymemo          = db2HTML(rsget("CS_MSG"))
                FItemList(i).FSellPrice             = rsget("SALE_PRICE")
                FItemList(i).FRealSellPrice         = rsget("SALE_PRICE")
                FItemList(i).FItemOrderCount        = rsget("QTY")
                FItemList(i).FrequireDetail         = ""
                FItemList(i).ForderDlvPay           = rsget("DELY_COST")
                FItemList(i).FSelldate              = rsget("ORDER_DT")
                FItemList(i).Forderserial           = ""

                IF IsNULL(FItemList(i).FrequireDetail) then FItemList(i).FrequireDetail=""

                FItemList(i).FoptionCnt             = 0 ''rsget("optionCnt")
                FItemList(i).FDuppExists            = 0 ''rsget("DuppExists")
                FItemList(i).FOrgDetailKey          = rsget("ORDER_SEQ")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
    end function

	public Function IsAllMatched(iOutMallOrderSerial)
	    dim i
	    For i=LBound(FItemList) to UBound(FItemList)
	        if IsObject(FItemList(i)) then
	            if FItemList(i).FOutMallOrderSerial=iOutMallOrderSerial then
	                IF (Not FItemList(i).IsItemMatched) then
	                    IsAllMatched = false
	                    Exit function
	                End IF
	            end if
	        end if
	    Next

	    IsAllMatched = true
    end function

    public function getDlvPayBySubPrice(isitename)
        dim ret : ret = 0
        dim ttlprice : ttlprice=0
        dim i

        ''if (isitename="dnshop") or (isitename="lotteCom") or (isitename="lotteimall") or (isitename="bandinlunis") then
            For i=LBound(FItemList) to UBound(FItemList)
    	        if IsObject(FItemList(i)) then
    	            ttlprice = ttlprice + FItemList(i).FRealSellPrice*FItemList(i).FItemOrderCount
    	        end if
    	    Next

    	    IF (isitename="dnshop") THEN
    	        if (ttlprice<50000) then ret=2000
    	    ELSEIF (isitename="lotteCom") then
    	        if (ttlprice<50000) then ret=3000
    	    ELSEIF (isitename="lotteimall") then
    	        if (ttlprice<50000) then ret=3000
    	    ELSEIF (isitename="cjmallITS") OR (isitename="itsCjmall") then
    	        if (ttlprice<50000) then ret=2500
    	    ELSEIF (isitename="bandinlunis") then
    	        if (ttlprice<30000) then ret=2500
    	    ELSEIF (isitename="fashionplus") then
    	        if (ttlprice<30000) then ret=2500
 			ELSEIF (isitename="ssg") then
    	        if (ttlprice<50000) then ret=3000
    	    ELSEIF (isitename="its29cm") then
    	        if (ttlprice<50000) then ret=2500
    	    ELSEIF (isitename="byulshopITS") THEN
    	        if (ttlprice<50000) then ret=2500
    	    ELSEIF (isitename="hiphoper") then
    	        if (ttlprice<50000) then ret=3000		'/아이띵소에서 힙합퍼쪽에 5만원 이하인 상품은 3천원 추가했다고함..
    	    ELSEIF (isitename="ithinksoshop") then
    	        if (ttlprice<10000) then ret=2500
    	    ELSEIF (isitename="wemakeprice") then
    	        if (ttlprice<9700) then ret=2500
    	    ELSEIF (isitename="itsKaKaoMakers") THEN
    	        if (ttlprice<50000) then ret=2500
    	    ELSEIF (isitename="itsWadiz") THEN
    	        if (ttlprice<50000) then ret=2500
    	    ELSEIF (isitename="hmall1010") THEN
    	        if (ttlprice<50000) then ret=3000		'2020-01-23 14:03 김진영 | 2500 -> 3000으로 수정
    	    ELSE
    	        if (ttlprice<30000) then ret=2500
    	    end if
        ''end if
        getDlvPayBySubPrice = ret
    end function

    public Function getOneTmpOrder()
        Dim sqlStr

        sqlStr = "select T.* "
        sqlStr = sqlStr&" , i.itemname as matchItemName"
		sqlStr = sqlStr&" , IsNULL(o.optionname,'') as matchItemOptionName"
		sqlStr = sqlStr&" from db_temp.dbo.tbl_xSite_tmpOrder T"
        sqlStr = sqlStr&"   left join db_item.dbo.tbl_item i"
		sqlStr = sqlStr&"   on T.matchItemid=i.itemid"
		sqlStr = sqlStr&"   left join db_item.dbo.tbl_item_option o"
		sqlStr = sqlStr&"   on T.matchItemid=o.itemid"
		sqlStr = sqlStr&"   and T.matchItemoption=o.itemoption"
		sqlStr = sqlStr&" where OutMallOrderseq="&FRectOutMallOrderSeq

		'response.write sqlstr & "<Br>"
        rsget.Open sqlStr,dbget,1

        FResultCount = rsget.RecordCount
        FTotalCount  = FResultCount
        if  not rsget.EOF  then
            set FOneItem = new CxSiteTempOrderItem

			FOneItem.FOutMallOrderSeq		= rsget("OutMallOrderSeq")
			''FOneItem.Fcompanyid				= rsget("companyid")
			FOneItem.FSellSite				= rsget("SellSite")
			FOneItem.FSellSiteName	        = rsget("SellSiteName")
            FOneItem.FmatchItemID           = rsget("matchItemID")
            FOneItem.FmatchItemOption       = rsget("matchItemOption")
			FOneItem.FmatchItemName			= rsget("matchItemName")
			FOneItem.FmatchItemOptionName	= rsget("matchItemOptionName")
            FOneItem.ForderItemID			= rsget("orderItemID")
			FOneItem.ForderItemName			= rsget("orderItemName")
			FOneItem.ForderItemOption		= rsget("orderItemOption")
			FOneItem.ForderItemOptionName	= rsget("orderItemOptionName")
			FOneItem.FOutMallOrderSerial	= rsget("OutMallOrderSerial")
			FOneItem.FfoundPrdcode			= rsget("Prdcode")
'			FOneItem.Ftplprdcode			= rsget("tplprdcode")
			FOneItem.Ftplorderserial		= rsget("orderserial")
			FOneItem.FmatchState			= rsget("matchState")
            FOneItem.FOrderName             = db2HTML(rsget("OrderName"))
            FOneItem.FOrderTelNo            = db2HTML(rsget("OrderTelNo"))
            FOneItem.FOrderHpNo             = db2HTML(rsget("OrderHpNo"))
            FOneItem.FReceiveName           = db2HTML(rsget("ReceiveName"))
            FOneItem.FReceiveTelNo          = db2HTML(rsget("ReceiveTelNo"))
            FOneItem.FReceiveHpNo           = db2HTML(rsget("ReceiveHpNo"))
            FOneItem.FReceiveZipCode        = db2HTML(rsget("ReceiveZipCode"))
            FOneItem.FReceiveAddr1          = db2HTML(rsget("ReceiveAddr1"))
            FOneItem.FReceiveAddr2          = db2HTML(rsget("ReceiveAddr2"))
            FOneItem.Fdeliverymemo          = db2HTML(rsget("deliverymemo"))
            FOneItem.FSellPrice             = rsget("SellPrice")
            FOneItem.FRealSellPrice         = rsget("RealSellPrice")
            FOneItem.FItemOrderCount        = rsget("ItemOrderCount")
            FOneItem.FrequireDetail         = rsget("requireDetail")
            FOneItem.ForderDlvPay           = rsget("orderDlvPay")
            FOneItem.FSelldate              = rsget("Selldate")
            FOneItem.Forderserial           = rsget("orderserial")

            IF IsNULL(FOneItem.FrequireDetail) then FOneItem.FrequireDetail=""

            FOneItem.FRef_OutMallOrderSerial = rsget("Ref_OutMallOrderSerial")

            ''FOneItem.FoptionCnt             = rsget("optionCnt")
            ''FOneItem.FDuppExists            = rsget("DuppExists")
            ''FOneItem.FOrgDetailKey          = rsget("OrgDetailKey")
            ''FOneItem.FordercsGbn            = rsget("ordercsGbn")

        end if
		rsget.close
    end function

	'2017-01-10 18:35 김진영 새 펑션 추가
	public Function getOnlineTmpOrderRealInputList()
	    Dim sqlStr, paramInfo, i
	    paramInfo = Array(Array("@RETURN_VALUE",adInteger,adParamReturnValue,,0) _
    			,Array("@SellSite"	    , adVarchar	, adParamInput	,32, FRectSellSite) _
    			,Array("@OutMallOrderSerial" , adVarchar	, adParamInput	, 32 , FRectOutMallOrderSerial) _
    		)
		sqlStr = "db_temp.dbo.sp_TEN_xSiteTmpOrderRealInputList"
'rw sqlStr
        Call fnExecSPReturnRSOutput(sqlStr,paramInfo)

        FTotalCount = GetValue(paramInfo, "@RETURN_VALUE")
        FtotalPage  = Int ( (FTotalCount - 1) / FPageSize ) + 1
		If FTotalCount = 0 Then	FtotalPage = 1

        FResultCount = rsget.RecordCount
        redim preserve FItemList(FResultCount)

        if  not rsget.EOF  then
		do Until rsget.Eof

			set FItemList(i) = new CxSiteTempOrderItem

			FItemList(i).forderemail		= rsget("orderemail")
			FItemList(i).fcountryCode		= rsget("countryCode")
			FItemList(i).FOutMallOrderSeq		= rsget("OutMallOrderSeq")
			''FItemList(i).Fcompanyid				= rsget("companyid")
			FItemList(i).FSellSite				= rsget("SellSite")
			FItemList(i).FSellSiteName	        = rsget("SellSiteName")
            FItemList(i).FmatchItemID           = rsget("matchItemID")
            FItemList(i).FmatchItemOption       = rsget("matchItemOption")
			FItemList(i).FmatchItemName			= rsget("matchItemName")
			FItemList(i).FmatchItemOptionName	= rsget("matchItemOptionName")
            FItemList(i).ForderItemID			= rsget("orderItemID")
			FItemList(i).ForderItemName			= rsget("orderItemName")
			FItemList(i).ForderItemOption		= rsget("orderItemOption")
			FItemList(i).ForderItemOptionName	= rsget("orderItemOptionName")
			FItemList(i).FOutMallOrderSerial	= rsget("OutMallOrderSerial")
			FItemList(i).FfoundPrdcode			= rsget("foundPrdcode")
'			FItemList(i).Ftplprdcode			= rsget("tplprdcode")
			FItemList(i).Ftplorderserial		= rsget("tplorderserial")
			FItemList(i).FmatchState			= rsget("matchState")
            FItemList(i).FOrderName             = db2HTML(rsget("OrderName"))
            FItemList(i).FOrderTelNo            = db2HTML(rsget("OrderTelNo"))
            FItemList(i).FOrderHpNo             = db2HTML(rsget("OrderHpNo"))
            FItemList(i).FReceiveName           = db2HTML(rsget("ReceiveName"))
            FItemList(i).FReceiveTelNo          = db2HTML(rsget("ReceiveTelNo"))
            FItemList(i).FReceiveHpNo           = db2HTML(rsget("ReceiveHpNo"))
            FItemList(i).FReceiveZipCode        = db2HTML(rsget("ReceiveZipCode"))
            FItemList(i).FReceiveAddr1          = db2HTML(rsget("ReceiveAddr1"))
            FItemList(i).FReceiveAddr2          = db2HTML(rsget("ReceiveAddr2"))
            FItemList(i).Fdeliverymemo          = db2HTML(rsget("deliverymemo"))
            FItemList(i).FSellPrice             = rsget("SellPrice")
            FItemList(i).FRealSellPrice         = rsget("RealSellPrice")
            FItemList(i).FItemOrderCount        = rsget("ItemOrderCount")
            FItemList(i).FrequireDetail         = rsget("requireDetail")
            FItemList(i).ForderDlvPay           = rsget("orderDlvPay")
            FItemList(i).FSelldate              = rsget("Selldate")
            FItemList(i).Forderserial           = rsget("orderserial")

            IF IsNULL(FItemList(i).FrequireDetail) then FItemList(i).FrequireDetail=""

            FItemList(i).FoptionCnt             = rsget("optionCnt")
            FItemList(i).FDuppExists            = rsget("DuppExists")
            FItemList(i).FOrgDetailKey          = rsget("OrgDetailKey")
            FItemList(i).FordercsGbn            = rsget("ordercsGbn")
            FItemList(i).Fbeadaldiv             = rsget("beadaldiv")
            FItemList(i).FItemdiv          		= rsget("itemdiv")
			FItemList(i).FPaydate          		= rsget("paydate")
			i=i+1
			rsget.movenext
		loop
        end if
		rsget.close
    end Function

	public Function getOnlineTmpOrderList(byval isLike)
	    Dim sqlStr, paramInfo, i
		if (FRectInc3pl="") then FRectInc3pl=NULL

	    paramInfo = Array(Array("@RETURN_VALUE",adInteger,adParamReturnValue,,0) _
    			,Array("@CurrPage"		, adInteger	, adParamInput	,		, FCurrPage)	_
    			,Array("@PageSize"		, adInteger	, adParamInput	,		, FPageSize) _
    			,Array("@SellSite"	    , adVarchar	, adParamInput	,32, FRectSellSite) _
    			,Array("@RectMatchState" , adVarchar	, adParamInput	, 10 , FRectMatchState) _
    			,Array("@OutMallOrderSerial" , adVarchar	, adParamInput	, 32 , FRectOutMallOrderSerial) _
    			,Array("@OrderSerial" , adVarchar	, adParamInput	, 16 , FRectOrderSerial) _
    			,Array("@RectCsViewYn" , adVarchar	, adParamInput	, 10 , FRectCsViewYn) _
				,Array("@regYYYYMMDD" , adVarchar	, adParamInput	, 10 , FRectregYYYYMMDD) _
				,Array("@is3plonly" , adInteger	, adParamInput	,  , FRectInc3pl) _
				,Array("@RectOverseaViewYn" , adVarchar	, adParamInput	, 10 , FRectOverseaViewYn) _
    		)

    	IF (isLike) then
    	    sqlStr = "db_temp.dbo.sp_TEN_xSiteTmpOrderListSearch"
    	ELSE
            sqlStr = "db_temp.dbo.sp_TEN_xSiteTmpOrderList"
        end if
''rw sqlStr &FCurrPage&","&FPageSize&","&FRectSellSite&","&FRectMatchState&","&FRectOutMallOrderSerial&","&FRectOrderSerial&","&FRectCsViewYn
        Call fnExecSPReturnRSOutput(sqlStr,paramInfo)

        FTotalCount = GetValue(paramInfo, "@RETURN_VALUE")
        FtotalPage  = Int ( (FTotalCount - 1) / FPageSize ) + 1
		If FTotalCount = 0 Then	FtotalPage = 1

        FResultCount = rsget.RecordCount
        redim preserve FItemList(FResultCount)

        if  not rsget.EOF  then
		do Until rsget.Eof

			set FItemList(i) = new CxSiteTempOrderItem

			FItemList(i).forderemail		= rsget("orderemail")
			FItemList(i).fcountryCode		= rsget("countryCode")
			FItemList(i).FOutMallOrderSeq		= rsget("OutMallOrderSeq")
			''FItemList(i).Fcompanyid				= rsget("companyid")
			FItemList(i).FSellSite				= rsget("SellSite")
			FItemList(i).FSellSiteName	        = rsget("SellSiteName")
            FItemList(i).FmatchItemID           = rsget("matchItemID")
            FItemList(i).FmatchItemOption       = rsget("matchItemOption")
			FItemList(i).FmatchItemName			= rsget("matchItemName")
			FItemList(i).FmatchItemOptionName	= rsget("matchItemOptionName")
            FItemList(i).ForderItemID			= rsget("orderItemID")
			FItemList(i).ForderItemName			= rsget("orderItemName")
			FItemList(i).ForderItemOption		= rsget("orderItemOption")
			FItemList(i).ForderItemOptionName	= rsget("orderItemOptionName")
			FItemList(i).FOutMallOrderSerial	= rsget("OutMallOrderSerial")
			FItemList(i).FfoundPrdcode			= rsget("foundPrdcode")
'			FItemList(i).Ftplprdcode			= rsget("tplprdcode")
			FItemList(i).Ftplorderserial		= rsget("tplorderserial")
			FItemList(i).FmatchState			= rsget("matchState")
            FItemList(i).FOrderName             = db2HTML(rsget("OrderName"))
            FItemList(i).FOrderTelNo            = db2HTML(rsget("OrderTelNo"))
            FItemList(i).FOrderHpNo             = db2HTML(rsget("OrderHpNo"))
            FItemList(i).FReceiveName           = db2HTML(rsget("ReceiveName"))
            FItemList(i).FReceiveTelNo          = db2HTML(rsget("ReceiveTelNo"))
            FItemList(i).FReceiveHpNo           = db2HTML(rsget("ReceiveHpNo"))
            FItemList(i).FReceiveZipCode        = db2HTML(rsget("ReceiveZipCode"))
            FItemList(i).FReceiveAddr1          = db2HTML(rsget("ReceiveAddr1"))
            FItemList(i).FReceiveAddr2          = db2HTML(rsget("ReceiveAddr2"))
            FItemList(i).Fdeliverymemo          = db2HTML(rsget("deliverymemo"))
            FItemList(i).FSellPrice             = rsget("SellPrice")
            FItemList(i).FRealSellPrice         = rsget("RealSellPrice")
            FItemList(i).FItemOrderCount        = rsget("ItemOrderCount")
            FItemList(i).FrequireDetail         = rsget("requireDetail")
            FItemList(i).ForderDlvPay           = rsget("orderDlvPay")
            FItemList(i).FSelldate              = rsget("Selldate")
            FItemList(i).Forderserial           = rsget("orderserial")

            IF IsNULL(FItemList(i).FrequireDetail) then FItemList(i).FrequireDetail=""

            FItemList(i).FoptionCnt             = rsget("optionCnt")
            FItemList(i).FDuppExists            = rsget("DuppExists")
            FItemList(i).FOrgDetailKey          = rsget("OrgDetailKey")
            FItemList(i).FordercsGbn            = rsget("ordercsGbn")
            FItemList(i).Fbeadaldiv             = rsget("beadaldiv")
            FItemList(i).FItemdiv             = rsget("itemdiv")
			FItemList(i).FFFExists             = rsget("FFExists")
            IF (isLike) then
                FItemList(i).FaddDlvExists            = rsget("addDlvExists")
                FItemList(i).FRef_OutMallOrderSerial = rsget("Ref_OutMallOrderSerial")

                FItemList(i).Fsellyn       = rsget("sellyn")
                FItemList(i).Flimityn      = rsget("limityn")
                FItemList(i).Flimitno      = rsget("limitno")
                FItemList(i).Flimitsold    = rsget("limitsold")
                FItemList(i).FSellcash     = rsget("Sellcash")
                FItemList(i).Foptusing     = rsget("optusing")
                FItemList(i).Foptsellyn    = rsget("optsellyn")
                FItemList(i).Foptlimityn   = rsget("optlimityn")
                FItemList(i).Foptlimitno   = rsget("optlimitno")
                FItemList(i).Foptlimitsold = rsget("optlimitsold")
                FItemList(i).Foptaddprice  = rsget("optaddprice")
            End IF
			FItemList(i).FShopifyOrderName             = rsget("shopifyOrderName")
			FItemList(i).FPaydate          		= rsget("paydate")
			FItemList(i).FOverseasPrice     	= rsget("overseasPrice")
			i=i+1
			rsget.movenext
		loop
        end if
		rsget.close
    end Function

    public Function fnOutmallOrderGetList
		fnOutmallOrderGetList =  clsConnDB.fnExecSPReturnRS("db_agirlOrder.dbo.[usp_Back_OutMallOrder_GetList]("&FRectSellSite&","&FOrderStatus&",'"&FSDate&"','"&FEDate&"','"&FIsMatching&"')")
	End Function

	public Function fnOutmallOrderGetDetail
		fnOutmallOrderGetDetail =  clsConnDB.fnExecSPReturnRS("db_agirlOrder.dbo.[usp_Back_OutMallOrder_GetDetailList]("&FSellSite&",'"&FOutMallOrderSerial&"')")
	End Function

    Private Sub Class_Initialize()
		redim  FItemList(0)

		FCurrPage =1
		FPageSize = 12
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
		FTotalPage =0

		FEMSPrice = 0
	End Sub

	Private Sub Class_Terminate()
	End Sub

	public Function HasPreScroll()
		HasPreScroll = StartScrollPage > 1
	end Function

	public Function HasNextScroll()
		HasNextScroll = FTotalPage > StartScrollPage + FScrollCount -1
	end Function

	public Function StartScrollPage()
		StartScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
	end Function
End Class

''EMS
function getEmsItemUsDollar(orderserial)
    dim orgItemprice : orgItemprice = GetTotalItemOrgPrice(orderserial)
    dim exchangeRate
    dim sqlStr
    sqlStr = "exec db_order.dbo.sp_Ten_Ems_exchangeRate 'USD'"

    rsget.CursorLocation = adUseClient
	rsget.CursorType = adOpenStatic
	rsget.LockType = adLockOptimistic
	rsget.Open sqlStr,dbget

	if Not rsget.Eof then
	    exchangeRate = rsget("exchangeRate")

	    if (exchangeRate>0) then
	        getEmsItemUsDollar = CLNG(orgItemprice/exchangeRate)
	    else
	        getEmsItemUsDollar = 0
	    end if
	else
	    getEmsItemUsDollar = 0
	end if

	rsget.close
end function

function GetTotalItemOrgPrice(orderserial)
	dim re,i, query1

	if orderserial = "" then
		GetTotalItemOrgPrice = 0
		exit function
	end if

	query1 = "select"
	query1 = query1 + " sum(d.orgitemcost*d.itemno) as orgitemcost"
	query1 = query1 + " from db_order.dbo.tbl_order_master m"
	query1 = query1 + " join db_order.dbo.tbl_order_detail d"
	query1 = query1 + " 	on m.orderserial=d.orderserial"
	query1 = query1 + " where m.cancelyn='N' and d.cancelyn='N' and d.itemid<>0"
	query1 = query1 + " and m.orderserial='"&orderserial&"'"

	'response.write query1 &"<br>"
	rsget.Open query1,dbget,1
	if not rsget.EOF  then
		re = rsget("orgitemcost")
	else
		re = 0
	end if
	rsget.close

	GetTotalItemOrgPrice = re
end function

''EMS 상품구분
function getEmsItemGubunName()
    getEmsItemGubunName = "Gift"
end function

''EMS 내용품명
function getEmsGoodNames()
    getEmsGoodNames = "stationery"
end function

function getEmsBoxWeight()
    getEmsBoxWeight = 200
end function

'' EMS 무게
function getEmsTotalWeight(orderserial)
    dim i, retVal, query1

	if orderserial = "" then
		GetTotalItemOrgPrice = 0
		exit function
	end if

    retVal = 0
	query1 = "select"
	query1 = query1 + " sum(i.itemWeight*d.itemno) as itemWeight"
	query1 = query1 + " from db_order.dbo.tbl_order_master m"
	query1 = query1 + " join db_order.dbo.tbl_order_detail d"
	query1 = query1 + " 	on m.orderserial=d.orderserial"
	query1 = query1 + " join db_item.dbo.tbl_item i"
	query1 = query1 + " 	on d.itemid=i.itemid"
	query1 = query1 + " where m.cancelyn='N' and d.cancelyn='N' and d.itemid<>0"
	query1 = query1 + " and m.orderserial='"&orderserial&"'"

	'response.write query1 &"<br>"
	rsget.Open query1,dbget,1
	if not rsget.EOF  then
		retVal = rsget("itemWeight")
	end if
	rsget.close

	getEmsTotalWeight = retVal + getEmsBoxWeight
end function

 ''EMS 추가 보헙 필요 여부
function isEmsInsureRequire(orderserial)
    ''(기본 보험 금액 : 60000 + CLng(getEmsTotalWeight/1000*10)/10*6750
    ''=IF(MOD((B12-98000),98000)=0,1800+INT((B12-98000)/98000)*450,1800+(INT((B12-98000)/98000)+1)*450)
    if (GetTotalItemOrgPrice(orderserial)>(60000 + CLng(getEmsTotalWeight(orderserial)/1000*10)/10*6750)) then
        isEmsInsureRequire = true
    else
        isEmsInsureRequire = false
    end if
end function

''EMS 추가 보헙 금액
function getEmsInsurePrice(orderserial)
    dim orgItemprice

    if isEmsInsureRequire(orderserial) then
        orgItemprice = GetTotalItemOrgPrice(orderserial)

        if (orgItemprice>98000) then
            getEmsInsurePrice = CLng((orgItemprice-98000)\98000)*450
            if (((orgItemprice-98000)/98000)>((orgItemprice-98000)\98000)) then getEmsInsurePrice = getEmsInsurePrice + 450
            getEmsInsurePrice = getEmsInsurePrice + 1800
        else
            getEmsInsurePrice = 1800
        end if
    else
        getEmsInsurePrice = 0
    end if
end function

function getxSiteDuppReceiverCheck(byval outmallorderserial)
	dim sql

	if outmallorderserial = "" then Exit function

	sql = "select COUNT(*) as duppCNT" &VbCRLF
    sql = sql& " from (" &VbCRLF
    sql = sql& " select OrderName,ReceiveName" &VbCRLF
    sql = sql& " from db_temp.dbo.tbl_xSite_TMPOrder" &VbCRLF
    sql = sql& " where OutMallOrderSerial='"&outmallorderserial&"'" &VbCRLF
    sql = sql& " group by OrderName,ReceiveName" &VbCRLF
    sql = sql& " ) T " &VbCRLF

	'response.write sql &"<Br>"
	rsget.Open sql,dbget,1
	if Not(rsget.EOF or rsget.BOF) then
		getxSiteDuppReceiverCheck = rsget("duppCNT")
	else
		getxSiteDuppReceiverCheck = 0
	end if
	rsget.Close
end function
%>

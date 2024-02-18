<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" lang="ko" xml:lang="ko">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
<meta http-equiv="X-UA-Compatible" content="IE=edge" />
<title></title>
<link rel="stylesheet" type="text/css" href="/css/adminDefault.css" />
<link rel="stylesheet" type="text/css" href="/css/adminCommon.css" />
<style type="text/css">
html {overflow:auto;}
.popWinV17 {overflow:hidden; position:absolute; left:0; top:0; right:0; bottom:0; width:100%; height:100%; font-family:"malgun Gothic","�������", Dotum, "����", sans-serif;}
.popWinV17 h1 {height:40px; padding:12px 15px 0; color:#fff; font-size:17px; background:#c80a0a; border-bottom:1px solid #d80a0a}
.popWinV17 h2 {position:relative; padding:12px 15px; color:#333; font-size:12px; font-weight: bold; background-color:#444; border-top:1px solid #666; font-family:"malgun Gothic","�������", Dotum, "����", sans-serif; z-index:55; color:#fff;}
.popContainerV17 {position:absolute; left:0; top:40px; right:0; bottom:90px; width:100%; border-bottom:1px solid #ddd;}
.contL {position:relative; width:65%; height:100%; border-right:1px solid #ddd; z-index:10; overflow-y:auto;}
.contR {position:absolute; right:0; top:0; bottom:0; width:30%; height:100%; border-left:1px solid #ddd;}
.tbListWrap {position:relative; width:100%; height:100%;}
.tbDataList, .thDataList {display:table; width:100%;}
.tbDataList li, .thDataList li {display:table; width:100%; margin-top:-1px; border-top:1px solid #ddd; border-bottom:1px solid #ddd; }
.thDataList li {height:33px; background-color:#eaeaea; border-top:2px solid #ccc; font-weight:bold;}
.tbDataList li {background-color:#fff; z-index:100;}
.tbDataList li p, .thDataList li p {display:table-cell; padding:7px; text-align:center; vertical-align:middle; line-height:1.4;}
.thDataList li p {white-space:nowrap;}
.handling {background-color:rgba(42,42,57,0.2) !important; height:30px; border:none;}
#sortable li {cursor:move;}
.popBtnWrap {position:absolute; left:0; bottom:0; width:100%; height:60px; text-align:center;}
.textOverflow {width:100%; display:block; text-overflow:ellipsis; overflow:hidden; white-space:nowrap;}
.btnMove {position:absolute; left:67.5%; top:50%; width:40px; height:70px; margin-top:-35px; margin-left:-20px; padding:0; border:none; background:transparent url(/images/btn_move_arrow.png) no-repeat 50% 50%; z-index:1000; cursor:pointer;}
</style>
</head>
<body>
<!-- �˾� ������ : �ּ� 1100*750 -->
<div class="popWinV17">
	<h1>Unit �˻�</h1>
	<div class="popContainerV17">
		<div class="contL">
			<h2>Unit ����</h2>
			<div class="tab" style="margin:-1px 0 0 -1px;">
				<ul>
					<li class="col11 selected"><a href="#unitType01">��ǰ</a></li>
					<li class="col11 "><a href="#unitType02">�̺�Ʈ</a></li>
					<li class="col11 "><a href="#unitType03">������</a></li>
				</ul>
			</div>
			<!-- ��ǰ Tab -->
			<div id="unitType01" class="unitPannel">
				<div class="searchWrap" style="border-top:none;">
					<div class="search">
						<ul>
							<li>
								<label class="formTit">ī�װ� :</label>
								<select class="formSlt" id="deal" title="�ɼ� ����">
									<option>��ü</option>
									<option>����</option>
									<option>ī�װ�</option>
									<option>�귣��</option>
									<option>��ǰ</option>
									<option>Ű����</option>
								</select>
								<select class="formSlt" id="deal3" title="�ɼ� ����">
									<option>��ü</option>
									<option>����</option>
									<option>ī�װ�</option>
									<option>�귣��</option>
									<option>��ǰ</option>
									<option>Ű����</option>
								</select>
								<select class="formSlt" id="deal4" title="�ɼ� ����">
									<option>��ü</option>
									<option>����</option>
									<option>ī�װ�</option>
									<option>�귣��</option>
									<option>��ǰ</option>
									<option>Ű����</option>
								</select>
							</li>
						</ul>
					</div>
					<dfn class="line"></dfn>
					<div class="search">
						<ul>
							<li>
								<label class="formTit" for="schWord">�˻��� :</label>
								<input type="text" class="formTxt" id="schWord" style="width:400px" placeholder="��ǰID �Ǵ� ��ǰ���� �Է��Ͽ� �˻��ϼ���." />
							</li>
						</ul>
					</div>
					<input type="button" class="schBtn" value="�˻�" />
				</div>
				<div class="tbListWrap tMar15">
					<div class="ftLt lPad10">
						<select class="formSlt" id="deal" title="�ɼ� ����">
							<option>�Ż�ǰ��</option>
							<option>�α��</option>
						</select>
					</div>
					<div class="ftRt pad10">
						<span>�˻���� : <strong>999,999</strong></span> <span class="lMar10">������ : <strong>1 / 30,000</strong></span>
					</div>
					<ul class="thDataList">
						<li>
							<p class="cell05"><input type="checkbox" /></p>
							<p class="cell10">��ǰ ID</p>
							<p class="cell10">�̹���</p>
							<p>��ǰ��</p>
							<p class="cell10">����</p>
							<p class="cell10">��ü ID</p>
							<p class="cell10">�Ǹſ���</p>
						</li>
					</ul>
					<ul class="tbDataList">
						<li>
							<p class="cell05"><input type="checkbox" /></p>
							<p class="cell10">17026941</p>
							<p class="cell10"><img src="http://webimage.10x10.co.kr/image/small/171/S001719788.jpg" width="50" height="50" border="0" /></p>
							<p class="lt">[����������] ����÷��ο��ǽ�(12346) ����÷��ο��ǽ�(12346)����÷��ο��ǽ�(12346)</p>
							<p class="cell10">316,000</p>
							<p class="cell10">milliens</p>
							<p class="cell10">ȫ�浿</p>
						</li>
						<li>
							<p class="cell05"><input type="checkbox" /></p>
							<p class="cell10">17026942</p>
							<p class="cell10"><img src="http://webimage.10x10.co.kr/image/small/171/S001719788.jpg" width="50" height="50" border="0" /></p>
							<p class="lt">�λ꺣� ö����</p>
							<p class="cell10">316,000</p>
							<p class="cell10">milliens</p>
							<p class="cell10">�ȿ���</p>
						</li>
					</ul>
					<div class="ct tPad20 cBk1">
						<a href="">[prev]</a>
						<a href=""><span class="cRd1">[1]</span></a>
						<a href="">[2]</a>
						<a href="">[3]</a>
						<a href="">[4]</a>
						<a href="">[5]</a>
						<a href="">[6]</a>
						<a href="">[7]</a>
						<a href="">[8]</a>
						<a href="">[9]</a>
						<a href="">[10]</a>
						<a href="">[next]</a>
					</div>
				</div>
			</div>
			<!--// ��ǰ Tab -->
			<!-- �̺�Ʈ Tab -->
			<div id="unitType02" class="unitPannel" style="display:none;">
				<div class="searchWrap" style="border-top:none;">
					<div class="search">
						<ul>
							<li>
								<label class="formTit">�Ⱓ :</label>
								<select class="formSlt" title="�ɼ� ����">
									<option>������</option>
									<option>������</option>
								</select>
								<input type="text" class="formTxt" id="term1" style="width:100px" placeholder="������" />
								<input type="image" src="/images/admin_calendar.png" alt="�޷����� �˻�" />
								~
								<input type="text" class="formTxt" id="term2" style="width:100px" placeholder="������" />
								<input type="image" src="/images/admin_calendar.png" alt="�޷����� �˻�" />
							</li>
						</ul>
					</div>
					<dfn class="line"></dfn>
					<div class="search">
						<ul>
							<li>
								<p class="formTit">�̺�Ʈ ���� :</p>
								<select class="formSlt" id="deal" title="�ɼ� ����">
									<option>��ü</option>
									<option>��������</option>
								</select>
							</li>
							<li>
								<p class="formTit">ī�װ� :</p>
								<select class="formSlt" id="deal" title="�ɼ� ����">
									<option>��ü</option>
									<option>�����ι���</option>
									<option>������</option>
								</select>
							</li>
						</ul>
					</div>
					<dfn class="line"></dfn>
					<div class="search">
						<ul>
							<li>
								<label class="formTit" for="schWord">�˻��� :</label>
								<input type="text" class="formTxt" id="schWord" style="width:400px" placeholder="�̺�Ʈ�ڵ� �Ǵ� �̺�Ʈ���� �Է��Ͽ� �˻��ϼ���." />
							</li>
						</ul>
					</div>
					<input type="button" class="schBtn" value="�˻�" />
				</div>
				<div class="tbListWrap tMar15">
					<div class="rt pad10">
						<span>�˻���� : <strong>999,999</strong></span> <span class="lMar10">������ : <strong>1 / 30,000</strong></span>
					</div>
					<ul class="thDataList">
						<li>
							<p class="cell05"><input type="checkbox" /></p>
							<p class="cell12">�̺�Ʈ �ڵ�</p>
							<p class="cell12">�̺�Ʈ ����</p>
							<p class="cell12">���</p>
							<p>�̺�Ʈ��</p>
							<p class="cell12">ī�װ�</p>
							<p class="cell12">������</p>
							<p class="cell12">������</p>
						</li>
					</ul>
					<ul class="tbDataList">
						<li>
							<p class="cell05"><input type="checkbox" /></p>
							<p class="cell12">17026941</p>
							<p class="cell12">��������</p>
							<p class="cell12"><img src="http://webimage.10x10.co.kr/image/small/171/S001719788.jpg" width="50" height="50" border="0" /></p>
							<p class="lt">���ѵ��� X ��ƼŰ���ͷ� ������ 2��</p>
							<p class="cell12">�����ι���</p>
							<p class="cell12">2017-05-05</p>
							<p class="cell12">2017-05-30</p>
						</li>
						<li>
							<p class="cell05"><input type="checkbox" /></p>
							<p class="cell12">17026941</p>
							<p class="cell12">��������</p>
							<p class="cell12"><img src="http://webimage.10x10.co.kr/image/small/171/S001719788.jpg" width="50" height="50" border="0" /></p>
							<p class="lt">���ѵ��� X ��ƼŰ���ͷ� ������ 2��</p>
							<p class="cell12">�����ι���</p>
							<p class="cell12">2017-05-05</p>
							<p class="cell12">2017-05-30</p>
						</li>
						<li>
							<p class="cell05"><input type="checkbox" /></p>
							<p class="cell12">17026941</p>
							<p class="cell12">��������</p>
							<p class="cell12"><img src="http://webimage.10x10.co.kr/image/small/171/S001719788.jpg" width="50" height="50" border="0" /></p>
							<p class="lt">���ѵ��� X ��ƼŰ���ͷ� ������ 2��</p>
							<p class="cell12">�����ι���</p>
							<p class="cell12">2017-05-05</p>
							<p class="cell12">2017-05-30</p>
						</li>
						<li>
							<p class="cell05"><input type="checkbox" /></p>
							<p class="cell12">17026941</p>
							<p class="cell12">��������</p>
							<p class="cell12"><img src="http://webimage.10x10.co.kr/image/small/171/S001719788.jpg" width="50" height="50" border="0" /></p>
							<p class="lt">���ѵ��� X ��ƼŰ���ͷ� ������ 2��</p>
							<p class="cell12">�����ι���</p>
							<p class="cell12">2017-05-05</p>
							<p class="cell12">2017-05-30</p>
						</li>
						<li>
							<p class="cell05"><input type="checkbox" /></p>
							<p class="cell12">17026941</p>
							<p class="cell12">��������</p>
							<p class="cell12"><img src="http://webimage.10x10.co.kr/image/small/171/S001719788.jpg" width="50" height="50" border="0" /></p>
							<p class="lt">���ѵ��� X ��ƼŰ���ͷ� ������ 2��</p>
							<p class="cell12">�����ι���</p>
							<p class="cell12">2017-05-05</p>
							<p class="cell12">2017-05-30</p>
						</li>
					</ul>
					<div class="ct tPad20 bPad20 cBk1">
						<a href="">[prev]</a>
						<a href=""><span class="cRd1">[1]</span></a>
						<a href="">[2]</a>
						<a href="">[3]</a>
						<a href="">[4]</a>
						<a href="">[5]</a>
						<a href="">[6]</a>
						<a href="">[7]</a>
						<a href="">[8]</a>
						<a href="">[9]</a>
						<a href="">[10]</a>
						<a href="">[next]</a>
					</div>
				</div>
			</div>
			<!--// �̺�Ʈ Tab -->
			<!-- ������ Tab -->
			<div id="unitType03" class="unitPannel" style="display:none;">
				<div class="searchWrap" style="border-top:none;">
					<div class="search">
						<ul>
							<li>
								<label class="formTit">������ :</label>
								<input type="text" class="formTxt" id="term1" style="width:100px" placeholder="������" />
								<input type="image" src="/images/admin_calendar.png" alt="�޷����� �˻�" />
								~
								<input type="text" class="formTxt" id="term2" style="width:100px" placeholder="������" />
								<input type="image" src="/images/admin_calendar.png" alt="�޷����� �˻�" />
							</li>
						</ul>
					</div>
					<dfn class="line"></dfn>
					<div class="search">
						<ul>
							<li>
								<label class="formTit">ī�װ� :</label>
								<select class="formSlt" id="deal" title="�ɼ� ����">
									<option>��ü</option>
									<option>����</option>
									<option>ī�װ�</option>
									<option>�귣��</option>
									<option>��ǰ</option>
									<option>Ű����</option>
								</select>
								<select class="formSlt" id="deal3" title="�ɼ� ����">
									<option>��ü</option>
									<option>����</option>
									<option>ī�װ�</option>
									<option>�귣��</option>
									<option>��ǰ</option>
									<option>Ű����</option>
								</select>
							</li>
						</ul>
					</div>
					<dfn class="line"></dfn>
					<div class="search">
						<ul>
							<li>
								<label class="formTit" for="schWord">�˻��� :</label>
								<input type="text" class="formTxt" id="schWord" style="width:400px" placeholder="Ÿ��Ʋ�� �˻��ϼ���." />
							</li>
						</ul>
					</div>
					<input type="button" class="schBtn" value="�˻�" />
				</div>
				<div class="tbListWrap tMar15">
					<div class="rt pad10">
						<span>�˻���� : <strong>999,999</strong></span> <span class="lMar10">������ : <strong>1 / 30,000</strong></span>
					</div>
					<ul class="thDataList">
						<li>
							<p class="cell05"><input type="checkbox" /></p>
							<p class="cell05">Idx</p>
							<p class="cell12">ī�װ�1</p>
							<p class="cell15">ī�װ�2</p>
							<p class="cell12">�̹���</p>
							<p>Ÿ��Ʋ</p>
							<p class="cell12">������</p>
						</li>
					</ul>
					<ul class="tbDataList">
						<li>
							<p class="cell05"><input type="checkbox" /></p>
							<p class="cell05">17</p>
							<p class="cell12">��ġ����Ŀ</p>
							<p class="cell15">MOVIE</p>
							<p class="cell12"><img src="http://webimage.10x10.co.kr/image/small/171/S001719788.jpg" width="50" height="50" border="0" /></p>
							<p class="lt">vol.63 &lt;�αٵα� ������&gt;</p>
							<p class="cell12">2017-05-30</p>
						</li>
						<li>
							<p class="cell05"><input type="checkbox" /></p>
							<p class="cell05">17</p>
							<p class="cell12">�÷���</p>
							<p class="cell15">TALK &gt; AZIT&</p>
							<p class="cell12"><img src="http://webimage.10x10.co.kr/image/small/171/S001719788.jpg" width="50" height="50" border="0" /></p>
							<p class="lt">������ �ٳ���� ȫ�� ��ȭ ����!</p>
							<p class="cell12">2017-05-30</p>
						</li>
						<li>
							<p class="cell05"><input type="checkbox" /></p>
							<p class="cell05">17</p>
							<p class="cell12">��ġ����Ŀ</p>
							<p class="cell15">!NSPIRATION &gt; DESIGN</p>
							<p class="cell12"><img src="http://webimage.10x10.co.kr/image/small/171/S001719788.jpg" width="50" height="50" border="0" /></p>
							<p class="lt">vol.63 &lt;�αٵα� ������&gt;</p>
							<p class="cell12">2017-05-30</p>
						</li>
						<li>
							<p class="cell05"><input type="checkbox" /></p>
							<p class="cell05">17</p>
							<p class="cell12">��ġ����Ŀ</p>
							<p class="cell15">THING. &gt; thingthing</p>
							<p class="cell12"><img src="http://webimage.10x10.co.kr/image/small/171/S001719788.jpg" width="50" height="50" border="0" /></p>
							<p class="lt">vol.63 &lt;�αٵα� ������&gt;</p>
							<p class="cell12">2017-05-30</p>
						</li>
					</ul>
					<div class="ct tPad20 bPad20 cBk1">
						<a href="">[prev]</a>
						<a href=""><span class="cRd1">[1]</span></a>
						<a href="">[2]</a>
						<a href="">[3]</a>
						<a href="">[4]</a>
						<a href="">[5]</a>
						<a href="">[6]</a>
						<a href="">[7]</a>
						<a href="">[8]</a>
						<a href="">[9]</a>
						<a href="">[10]</a>
						<a href="">[next]</a>
					</div>
				</div>
			</div>
			<!--// ������ Tab -->
		</div>

		<input type="button" class="btnMove" title="�����ؼ� ���" />

		<div class="contR">
			<h2 style="margin-left:-1px;">Unit ���� ����</h2>
			<div class="tbListWrap">
				<ul class="thDataList">
					<li>
						<p class="cell10"><input type="checkbox" /></p>
						<p class="cell25">Unit ����</p>
						<p>Unit��</p>
					</li>
				</ul>
				<ul id="sortable" class="tbDataList">
					<li>
						<p class="cell10"><input type="checkbox" /></p>
						<p class="cell25">��ǰ</p>
						<p class="lt"><span class="textOverflow">[����������] ����÷��ο��ǽ�(12346)</span></p>
					</li>
					<li>
						<p class="cell10"><input type="checkbox" /></p>
						<p class="cell25">������</p>
						<p class="lt"><span class="textOverflow">sunny tote bag yellow</span></p>
					</li>
					<li>
						<p class="cell10"><input type="checkbox" /></p>
						<p class="cell25">�̺�Ʈ</p>
						<p class="lt"><span class="textOverflow">�λ꺣� ö����</span></p>
					</li>
					<li>
						<p class="cell10"><input type="checkbox" /></p>
						<p class="cell25">������</p>
						<p class="lt"><span class="textOverflow">sunny tote bag yellow</span></p>
					</li>
					<li>
						<p class="cell10"><input type="checkbox" /></p>
						<p class="cell25">�̺�Ʈ</p>
						<p class="lt"><span class="textOverflow">�λ꺣� ö����</span></p>
					</li>
				</ul>
				<div class="pad10 rt">
					<input type="button" class="btn" value="���û���" onclick="" />
				</div>
			</div>
		</div>
	</div>
	<div class="popBtnWrap">
		<input type="button" value="���ÿϷ�" onclick="" class="cRd1" style="width:100px; height:30px;" />
		<input type="button" value="���" onclick="" style="width:100px; height:30px;" />
	</div>
</div>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/jquery-ui-1.10.3.custom.min.js"></script>
<script>
$(function() {
	$("#sortable").sortable({
		placeholder:"handling"
	}).disableSelection();

	$(".tab li").click(function() {
		$(".tab li").removeClass('selected');
		$(this).addClass('selected');
		$('.unitPannel').hide();
		var activeTab = $(this).find("a").attr("href");
		$(activeTab).show();
		return false;
	});
});
</script>
</body>
</html>
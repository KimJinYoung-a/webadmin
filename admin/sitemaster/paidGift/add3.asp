<%@ language=vbscript %>
    <% option explicit %>
        <% '#############################################################
' Description : ���� ����ǰ ���� '	History		: 2022.01.05 ������ ����
' ############################################################# %>
            <!-- #include virtual="/lib/function.asp"-->
            <!-- #include virtual="/lib/db/dbopen.asp" -->
            <!-- #include virtual="/lib/util/htmllib.asp" -->
            <!-- #include virtual="/admin/incSessionAdmin.asp" -->
            <!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
            </p>
            <link rel="stylesheet" type="text/css" href="/css/commonV20.css" />
            <div>
                <!-- �������� ����� -->
                <div class="content paidGift add">
                    <div class="paidGift_top">
                        <div class="search_wrap">
                            <select name="" id="select01">
                                <option value="">����</option>
                                <option value="">�÷�������</option>
                                <option value="">�������ǰ</option>
                            </select>
                            <select name="" id="select02">
                                <option value="">�������</option>
                                <option value="">������</option>
                                <option value="">���࿹��</option>
                                <option value="">����</option>
                                <option value="">������</option>
                            </select>
                            <div class="input_wrap">
                                <select name="" id="select03">
                                    <option value="">���μ�</option>
                                    <option value="">�����</option>
                                </select>
                                <span></span>
                                <button class="btn_select">�����ϱ�</button>
                                <li class="selected" style="display:none;">������<button class="close"><img src="https://fiximage.10x10.co.kr/web2019/diary2020/ico_close.png" alt=""></button></li>
                            </div>
                             <div class="input_wrap">
                                <select name="" id="select04">
                                    <option value="">��������</option>
                                    <option value="">���� ù ����</option>
                                    <option value="">�ֱ� �����ݾ�</option>
                                    <option value="">ȸ�����</option>
                                    <option value="">���űݾ�</option>
                                    <option value="">����Ƚ��</option>
                                    <option value="">��ǰ</option>
                                    <option value="">ī�װ�</option>
                                    <option value="">�귣��</option>
                                    <option value="">��ȹ��/�̺�Ʈ</option>
                                </select>
                                <li class="selected" style="display:none;">��ȹ��/�̺�Ʈ<button class="close"><img src="https://fiximage.10x10.co.kr/web2019/diary2020/ico_close.png" alt=""></button></li>
                                <li class="selected" style="display:none;">��ǰ<button class="close"><img src="https://fiximage.10x10.co.kr/web2019/diary2020/ico_close.png" alt=""></button></li>
                            </div>
                            <div class="input_wrap">
                                <select name="" id="select05">
                                    <option value="">��������/��ǰ/����ǰ��</option>
                                    <option value="">�������� ��ȣ</option>
                                    <option value="">����ǰ�ڵ�</option>
                                    <option value="">��ǰ�ڵ�</option>
                                </select>
                                <span></span>
                                <input type="text" placeholder="�˻�� �Է����ּ���">
                            </div>
                            <button class="btn_search"><img src="https://webadmin.10x10.co.kr/images/icon/search.png">�˻��ϱ�</button>
                        </div>
                        <div class="tgl_wrap">
                            <span>���� ����� �������ø� ����</span>
                            <div class="tgl_btn">
                                <input type="checkbox" id="tgl_btn_my">
                                <label for="tgl_btn_my"></label>
                            </div>
                        </div>
                    </div>
                    <div class="paidGift_aside">
                        <div class="list_wrap">
                            <div class="list_top">
                                <li>�� <span>103</span>��</li>
                            </div>
                            <div class="list_cont">
                                <a href=""><div class="cont cont_new">
                                    <span class="num">103</span>
                                    <ul>
                                        <h3>[����] �� ��������</h3>
                                        <li>2021.01.23 - 2022.01.27</li>
                                        <li>�������� : <span>���� ù ����</span></li>
                                        <li><span></span></li>
                                        <li class="state st01" style="display:none;"></li>
                                        <p>12:31:20 �ڵ�����</p>
                                    </ul>
                                    <button class="delete"><img src="https://webadmin.10x10.co.kr/images/icon/trash_red.png"></button>
                                </div></a>
                                <!-- �ӽ����� �������� -->
                                <a href=""><div class="cont" style="display:none;">
                                    <span class="num">103</span>
                                    <ul>
                                        <h3>[�÷�������] 1�� ù���� �÷�������</h3>
                                        <li>2021.01.23 - 2022.01.27</li>
                                        <li>�������� : <span>����A, ����B</span></li>
                                        <li>���μ� : <span>������</span></li>
                                        <li class="state st03">�ӽ�����</li>
                                    </ul>
                                </div></a>
                                <a href=""><div class="cont">
                                    <span class="num">103</span>
                                    <ul>
                                        <h3>[�÷�������] 1�� ù���� �÷�������</h3>
                                        <li>2021.01.23 - 2022.01.27</li>
                                        <li>�������� : <span>����A, ����B</span></li>
                                        <li>���μ� : <span>������</span></li>
                                        <li class="state st01">���¿���</li>
                                    </ul>
                                </div></a>
                                <a href=""><div class="cont">
                                    <span class="num">103</span>
                                    <ul>
                                        <h3>[�÷�������] 1�� ù���� �÷�������</h3>
                                        <li>2021.01.23 - 2022.01.27</li>
                                        <li>�������� : <span>����A, ����B</span></li>
                                        <li>���μ� : <span>������</span></li>
                                        <li class="state st02">������</li>
                                    </ul>
                                </div></a>
                                <a href=""><div class="cont">
                                    <span class="num">103</span>
                                    <ul>
                                        <h3>[�÷�������] 1�� ù���� �÷�������</h3>
                                        <li>2021.01.23 - 2022.01.27</li>
                                        <li>�������� : <span>����A, ����B</span></li>
                                        <li>���μ� : <span>������</span></li>
                                        <li class="state st02">������</li>
                                    </ul>
                                </div></a>
                                <a href=""><div class="cont">
                                    <span class="num">103</span>
                                    <ul>
                                        <h3>[�÷�������] 1�� ù���� �÷�������</h3>
                                        <li>2021.01.23 - 2022.01.27</li>
                                        <li>�������� : <span>����A, ����B</span></li>
                                        <li>���μ� : <span>������</span></li>
                                        <li class="state st02">������</li>
                                    </ul>
                                </div></a>
                                <a href=""><div class="cont">
                                    <span class="num">103</span>
                                    <ul>
                                        <h3>[�÷�������] 1�� ù���� �÷�������</h3>
                                        <li>2021.01.23 - 2022.01.27</li>
                                        <li>�������� : <span>����A, ����B</span></li>
                                        <li>���μ� : <span>������</span></li>
                                        <li class="state st01">���¿���</li>
                                    </ul>
                                </div></a>
                                <a href=""><div class="cont">
                                    <span class="num">103</span>
                                    <ul>
                                        <h3>[�÷�������] 1�� ù���� �÷�������</h3>
                                        <li>2021.01.23 - 2022.01.27</li>
                                        <li>�������� : <span>����A, ����B</span></li>
                                        <li>���μ� : <span>������</span></li>
                                        <li class="state st03">����</li>
                                    </ul>
                                </div></a>
                            </div>
                            <div class="list_bottom">
                                <ul class="pagination">
                                    <li class="on"><a>1</a></li>
                                    <li class=""><a>2</a></li>
                                    <li class=""><a>3</a></li>
                                    <li class=""><a>4</a></li>
                                    <li class=""><a>5</a></li> 
                                    <li class=""><a>></a></li>
                                    <li class=""><a>>></a></li>
                                </ul>
                            </div>
                        </div>
                    </div>
                    <div class="paidGift_section">
                    <!-- ���ǳ��� : on -->
                        <div class="steps">
                            <li class="step on">���� �������� ����</li>
                            <li class="step on"><span></span>�Ⱓ����</li>
                            <li class="step on"><span></span>���ü���</li>
                            <li class="step"><span></span>��������</li>
                        </div>
                        <div class="step_wrap step03">
                            <div class="step_noti on"><span>������ ������ � ������ �����ұ��?</span></div>
                            <div class="btn_group">
                                <button class="type01 step03_01 on"><span>�÷��� ����</span><p class="img"><img></p></button>
                                <button class="type01 step03_02"><span>�������ǰ</span><p class="img"><img></p></button>
                            </div>
                            <!-- �÷������� -->
                            <div class="step_cont step03_01">
                             <li>�����ڿ��� ������ ��ǰ�� ���� ������ �����մϴ�. ������ �������ּ���.</li>
                             <button class="btn_que"><img src="https://webadmin.10x10.co.kr/images/icon/question.png" alt=""></button>
                                <div class="step_detail">
                                    <div class="step_detail_list">
                                        <button class="type02 on">
                                            <h3>�׷���� ��ǰ ���</h3>
                                            <li>�׷� ���� ���� ��ǰ�� �����մϴ�.</li>
                                            <li class="check"><img src="https://webadmin.10x10.co.kr/images/icon/check.png" alt=""></li>
                                        </button>
                                        <button class="type02">
                                            <h3>�ܼ� �׷� ����</h3>
                                            <li>��ǰ�׷��� �з��ϰ�, �׷쿡 �̸��� ���� �� �ֽ��ϴ�.</li>
                                            <li class="check"><img src="https://webadmin.10x10.co.kr/images/icon/check.png" alt=""></li>
                                        </button>
                                        <button class="type02">
                                            <h3>�ݾ׺� �׷� ����</h3>
                                            <li>���űݾ׺��� ��ǰ�� �з��Ͽ�, ������ ������ ��쿡�� ���� �����ϰ� �մϴ�.</li>
                                            <li class="check"><img src="https://webadmin.10x10.co.kr/images/icon/check.png" alt=""></li>
                                        </button>
                                    </div>
                                </div>
                                <!-- �׷���� ��ǰ ��� -->
                                <div class="step_detail group group01_01" style="display:block;">
                                    <div class="table_wrap">
                                        <div class="ttop">
                                            <span>�� 0��</span><span class="selected">���õ� ��ǰ 1��</span>
                                            <button>���û�ǰ ����</button>
                                        </div>
                                        <div class="table">
                                            <ul class="thead">
                                                <li><input type="checkbox"></li>
                                                <li>��ǰ�ڵ�</li>
                                                <li>��ǥ�̹���/�귣��/��ǰ��</li>
                                                <li>�߰�����</li>
                                                <li>��� ������<span>��������/��ü����</span></li>
                                                <li>�ǸŰ�</li>
                                                <li>���԰�</li>
                                                <li>����</li>
                                                <li>��౸��</li>
                                                <li></li>
                                            </ul>
                                            <div class="tbody_wrap add">
                                                <ul class="tbody">
                                                    <a href=""><span>
                                                        + ��ǰ �߰��ϱ�
                                                    </span></a>
                                                </ul>
                                            </div>
                                            <div class="tbody_wrap">
                                                <ul class="tbody">
                                                    <li><input type="checkbox"></li>
                                                    <li>12345678</li>
                                                    <li>
                                                        <img class="prd_img">
                                                        <ul><p class="prd_brand">dailylike</p>
                                                        <p class="prd_name">�ų����� ����ī�� ��Ű��</p></ul>
                                                    </li>
                                                    <li>
                                                        <ul><p class="badge">����[MDƯ��]</p>
                                                        <p class="info">���ǻ���</p></ul>
                                                    </li>
                                                    <li>
                                                        <ul>
                                                            <p class="percent">99%</p>
                                                            <ul class="bar">
                                                                <p class="bar_gray"></p>
                                                                <p class="bar_black"></p>
                                                            </ul>
                                                            <p class="">99/100</p>
                                                        </ul>
                                                    </li>
                                                    <li>
                                                        <ul>
                                                            <p class="price p_type01"><span>38,000</span></p>
                                                            <p class="price p_type02">7% ����<span>38,000</span></p>
                                                            <p class="price p_type03">7% ����<span>38,000</span></p>
                                                            <p class="price p_type04">15% �÷�������<span>38,000</span></p>
                                                        </ul>
                                                    </li>
                                                    <li>
                                                        <ul>
                                                            <p class="p_type01">38,000</p>
                                                            <p class="p_type02">23,902</p>
                                                            <p class="p_type03">23,902</p>
                                                            <p class="p_type04">23,902</p>
                                                        </ul>
                                                    </li>
                                                    <li>
                                                        <ul>
                                                            <p class="p_type01">27%</p>
                                                            <p class="p_type02">26%</p>
                                                            <p class="p_type03">26%</p>
                                                            <p class="p_type04">28%</p>
                                                        </ul>
                                                    </li>
                                                    <li>
                                                        ����
                                                    </li>
                                                    <li>
                                                        <ul>
                                                            <button>����</button>
                                                            <button>����</button>
                                                        </ul>
                                                    </li>
                                                </ul>
                                            </div>
                                            <div class="tbody_wrap">
                                                <ul class="tbody">
                                                    <li><input type="checkbox"></li>
                                                    <li>12345678</li>
                                                    <li>
                                                        <img class="prd_img">
                                                        <ul><p class="prd_brand">dailylike</p>
                                                        <p class="prd_name">�ų����� ����ī�� ��Ű��</p></ul>
                                                    </li>
                                                    <li>
                                                        <ul><p class="badge"></p>
                                                        <p class="info"></p></ul>
                                                    </li>
                                                    <li>
                                                        <ul>
                                                            <p class="percent"></p>
                                                            <ul class="bar" style="display:none;">
                                                                <p class="bar_gray"></p>
                                                                <p class="bar_black"></p>
                                                            </ul>
                                                            <p class="">99</p>
                                                        </ul>
                                                    </li>
                                                    <li>
                                                        <ul>
                                                            <p class="price p_type01"><span>38,000</span></p>
                                                            <p class="price p_type02">7% ����<span>38,000</span></p>
                                                            <p class="price p_type03">7% ����<span>38,000</span></p>
                                                            <p class="price p_type04">15% �÷�������<span>38,000</span></p>
                                                        </ul>
                                                    </li>
                                                    <li>
                                                        <ul>
                                                            <p class="p_type01">38,000</p>
                                                            <p class="p_type02">23,902</p>
                                                            <p class="p_type03">23,902</p>
                                                            <p class="p_type04">23,902</p>
                                                        </ul>
                                                    </li>
                                                    <li>
                                                        <ul>
                                                            <p class="p_type01">27%</p>
                                                            <p class="p_type02">26%</p>
                                                            <p class="p_type03">26%</p>
                                                            <p class="p_type04">28%</p>
                                                        </ul>
                                                    </li>
                                                    <li>
                                                        ����
                                                    </li>
                                                    <li>
                                                        <ul>
                                                            <button>����</button>
                                                            <button>����</button>
                                                        </ul>
                                                    </li>
                                                </ul>
                                            </div>
                                        </div>
                                    </div>
                                </div>
                                <!-- �ܼ� �׷� ���� -->
                                 <div class="step_detail group group02_01" style="display:none;">
                                    <div class="tab">
                                        <button class="added on"><span class="sort"><img src="https://webadmin.10x10.co.kr/images/icon/drag.png"></span><li>���̾���� ���Ϲ����ص� ��ƼĿ��</li><img src="https://webadmin.10x10.co.kr/images/icon/edit.png" alt=""></button>
                                        <button class="added"><span class="sort"><img src="https://webadmin.10x10.co.kr/images/icon/drag.png"></span><li>ġƮŰ�� ���޸���</li><img src="https://webadmin.10x10.co.kr/images/icon/edit.png" alt=""></button>
                                        <button><li>+�׷� �߰��ϱ�</li></button>
                                    </div>
                                    <!-- �׷Ӿ��� ��ǰ ��ϰ� ����<div class="table_wrap"></div> -->
                                </div>
                                <!-- �ݾ׺� �׷� ���� -->
                                 <div class="step_detail group group03_01" style="display:none;">
                                    <div class="tab">
                                        <button class="added on"><li>10,000��~<span>&#65372;</span><span class="gr_cond">���űݾ�</span><span class="gr_name">10,000�� �̻� �����ߴٸ�</span></li><img src="https://webadmin.10x10.co.kr/images/icon/edit.png" alt=""></button>
                                        <button class="added"><li>10,000��~<span>&#65372;</span><span class="gr_cond">���űݾ�</span><span class="gr_name">10,000�� �̻� �����ߴٸ�</span></li><img src="https://webadmin.10x10.co.kr/images/icon/edit.png" alt=""></button>
                                        <button><li>+�ݾ׺� �׷� �߰��ϱ�</li></button>
                                    </div>
                                    <!-- �׷Ӿ��� ��ǰ ��ϰ� ����<div class="table_wrap"></div> -->
                                </div>
                            </div>
                            <!-- �������ǰ -->
                            <div class="step_cont step03_02 on">
                             <li>�����ڿ��� ������ ����ǰ�� ����� �����մϴ�. �����ڴ� �� �ϳ��� ����ǰ�� �����Ͽ� ���� �� �ֽ��ϴ�. ������ �������ּ���.</li>
                             <button class="btn_que"><img src="https://webadmin.10x10.co.kr/images/icon/question.png" alt=""></button>
                                <div class="step_detail">
                                    <div class="step_detail_list">
                                        <button class="type02 on">
                                            <h3>�׷���� ��ǰ ���</h3>
                                            <li>�׷� ���� ���� ��ǰ�� �����մϴ�.</li>
                                            <li class="check"><img src="https://webadmin.10x10.co.kr/images/icon/check.png" alt=""></li>
                                        </button>
                                        <button class="type02">
                                            <h3>�ܼ� �׷� ����</h3>
                                            <li>����ǰ�� �з��ϰ�, �׷쿡 �̸��� ���� �� �ֽ��ϴ�.</li>
                                            <li class="check"><img src="https://webadmin.10x10.co.kr/images/icon/check.png" alt=""></li>
                                        </button>
                                        <button class="type02">
                                            <h3>�ݾ׺� �׷� ����</h3>
                                            <li>���űݾ׺��� ����ǰ�� �з��Ͽ�, ������ ������ ��쿡�� ���� �����ϰ� �մϴ�.</li>
                                            <li class="check"><img src="https://webadmin.10x10.co.kr/images/icon/check.png" alt=""></li>
                                        </button>
                                    </div>
                                </div>
                                <!-- �׷���� ��ǰ ��� -->
                                <div class="step_detail group group01_02" style="display:block;">
                                    <div class="table_wrap">
                                        <div class="ttop">
                                            <span>�� 0��</span><span class="selected">���õ� ��ǰ 1��</span>
                                            <button>���û�ǰ ����</button>
                                        </div>
                                        <div class="table">
                                            <ul class="thead">
                                                <li><input type="checkbox"></li>
                                                <li>����ǰ�ڵ�</li>
                                                <li>��ǥ�̹���/�귣��/��ǰ��</li>
                                                <li>�߰�����</li>
                                                <li>��� ������<span>��������/��ü����</span></li>
                                                <li>�ǸŰ�</li>
                                                <li>���԰�</li>
                                                <li>����</li>
                                                <li>��౸��</li>
                                                <li></li>
                                            </ul>
                                            <div class="tbody_wrap add">
                                                <ul class="tbody">
                                                    <div>
                                                        + ����ǰ �߰��ϱ�
                                                        <div class="add_btn">
                                                            <button>����ǰ �űԵ��</button><span></span>
                                                            <button>���� ����ǰ �ҷ�����</button>
                                                        </div>
                                                    </div>
                                                </ul>
                                            </div>
                                            <div class="tbody_wrap">
                                                <ul class="tbody">
                                                    <li><input type="checkbox"></li>
                                                    <li>12345678</li>
                                                    <li>
                                                        <img class="prd_img">
                                                        <ul><p class="prd_brand">dailylike</p>
                                                        <p class="prd_name">�ų����� ����ī�� ��Ű��</p></ul>
                                                    </li>
                                                    <li>
                                                        <ul><p class="badge">����[MDƯ��]</p>
                                                        <p class="info">���ǻ���</p></ul>
                                                    </li>
                                                    <li>
                                                        <ul>
                                                            <p class="percent">99%</p>
                                                            <ul class="bar">
                                                                <p class="bar_gray"></p>
                                                                <p class="bar_black"></p>
                                                            </ul>
                                                            <p class="">99/100</p>
                                                        </ul>
                                                    </li>
                                                    <li>
                                                        <ul>
                                                            <p class="price p_type01"><span>38,000</span></p>
                                                            <p class="price p_type02">7% ����<span>38,000</span></p>
                                                            <p class="price p_type03">7% ����<span>38,000</span></p>
                                                            <p class="price p_type04">15% �÷�������<span>38,000</span></p>
                                                        </ul>
                                                    </li>
                                                    <li>
                                                        <ul>
                                                            <p class="p_type01">38,000</p>
                                                            <p class="p_type02">23,902</p>
                                                            <p class="p_type03">23,902</p>
                                                            <p class="p_type04">23,902</p>
                                                        </ul>
                                                    </li>
                                                    <li>
                                                        <ul>
                                                            <p class="p_type01">27%</p>
                                                            <p class="p_type02">26%</p>
                                                            <p class="p_type03">26%</p>
                                                            <p class="p_type04">28%</p>
                                                        </ul>
                                                    </li>
                                                    <li>
                                                        ����
                                                    </li>
                                                    <li>
                                                        <ul>
                                                            <button>����</button>
                                                            <button>����</button>
                                                        </ul>
                                                    </li>
                                                </ul>
                                            </div>
                                            <div class="tbody_wrap">
                                                <ul class="tbody">
                                                    <li><input type="checkbox"></li>
                                                    <li>12345678</li>
                                                    <li>
                                                        <img class="prd_img">
                                                        <ul><p class="prd_brand">dailylike</p>
                                                        <p class="prd_name">�ų����� ����ī�� ��Ű��</p></ul>
                                                    </li>
                                                    <li>
                                                        <ul><p class="badge"></p>
                                                        <p class="info"></p></ul>
                                                    </li>
                                                    <li>
                                                        <ul>
                                                            <p class="percent"></p>
                                                            <ul class="bar" style="display:none;">
                                                                <p class="bar_gray"></p>
                                                                <p class="bar_black"></p>
                                                            </ul>
                                                            <p class="">99</p>
                                                        </ul>
                                                    </li>
                                                    <li>
                                                        <ul>
                                                            <p class="price p_type01"><span>38,000</span></p>
                                                            <p class="price p_type02">7% ����<span>38,000</span></p>
                                                            <p class="price p_type03">7% ����<span>38,000</span></p>
                                                            <p class="price p_type04">15% �÷�������<span>38,000</span></p>
                                                        </ul>
                                                    </li>
                                                    <li>
                                                        <ul>
                                                            <p class="p_type01">38,000</p>
                                                            <p class="p_type02">23,902</p>
                                                            <p class="p_type03">23,902</p>
                                                            <p class="p_type04">23,902</p>
                                                        </ul>
                                                    </li>
                                                    <li>
                                                        <ul>
                                                            <p class="p_type01">27%</p>
                                                            <p class="p_type02">26%</p>
                                                            <p class="p_type03">26%</p>
                                                            <p class="p_type04">28%</p>
                                                        </ul>
                                                    </li>
                                                    <li>
                                                        ����
                                                    </li>
                                                    <li>
                                                        <ul>
                                                            <button>����</button>
                                                            <button>����</button>
                                                        </ul>
                                                    </li>
                                                </ul>
                                            </div>
                                        </div>
                                    </div>
                                </div>
                                <!-- �ܼ� �׷� ���� -->
                                 <div class="step_detail group group02_02" style="display:none;">
                                    <div class="tab">
                                        <button class="added on"><span class="sort"><img src="https://webadmin.10x10.co.kr/images/icon/drag.png"></span><li>���̾���� ���Ϲ����ص� ��ƼĿ��</li><img src="https://webadmin.10x10.co.kr/images/icon/edit.png" alt=""></button>
                                        <button class="added"><span class="sort"><img src="https://webadmin.10x10.co.kr/images/icon/drag.png"></span><li>ġƮŰ�� ���޸���</li><img src="https://webadmin.10x10.co.kr/images/icon/edit.png" alt=""></button>
                                        <button><li>+�׷� �߰��ϱ�</li></button>
                                    </div>
                                    <!-- �׷Ӿ��� ��ǰ ��ϰ� ����<div class="table_wrap"></div> -->
                                </div>
                                <!-- �ݾ׺� �׷� ���� -->
                                 <div class="step_detail group group03_02" style="display:none;">
                                    <div class="tab">
                                        <button class="added on"><li>10,000��~<span>&#65372;</span><span class="gr_cond">���űݾ�</span><span class="gr_name">10,000�� �̻� �����ߴٸ�</span></li><img src="https://webadmin.10x10.co.kr/images/icon/edit.png" alt=""></button>
                                        <button class="added"><li>10,000��~<span>&#65372;</span><span class="gr_cond">���űݾ�</span><span class="gr_name">10,000�� �̻� �����ߴٸ�</span></li><img src="https://webadmin.10x10.co.kr/images/icon/edit.png" alt=""></button>
                                        <button><li>+�ݾ׺� �׷� �߰��ϱ�</li></button>
                                    </div>
                                    <div class="table_wrap t02">
                                        <div class="ttop">
                                            <span>�� 0��</span><span class="selected">���õ� ��ǰ 1��</span>
                                            <button>���û�ǰ ����</button>
                                        </div>
                                        <div class="table">
                                            <ul class="thead">
                                                <li><input type="checkbox"></li>
                                                <li>��ǰ�ڵ�</li>
                                                <li>��ǥ�̹���/�귣��/��ǰ��</li>
                                                <li>�߰�����</li>
                                                <li>��� ������<span>��������/��ü����</span></li>
                                                <li>��� QR ����/���ϸ��� ����</li>
                                                <li></li>
                                            </ul>
                                            <!-- �׷� �߰� �� -->
                                            <div class="tbody_wrap none">
                                                <ul class="tbody">
                                                    �׷��� ���� �߰��ϸ� ����ǰ�� �߰��� �� �־��!
                                                </ul>
                                            </div>
                                            <div class="tbody_wrap add">
                                                <ul class="tbody">
                                                    <div>
                                                        + ����ǰ �߰��ϱ�
                                                        <div class="add_btn">
                                                            <button>����ǰ �űԵ��</button><span></span>
                                                            <button>���� ����ǰ �ҷ�����</button>
                                                        </div>
                                                    </div>
                                                </ul>
                                            </div>
                                            <div class="tbody_wrap">
                                                <ul class="tbody">
                                                    <li><input type="checkbox"></li>
                                                    <li>12345678</li>
                                                    <li>
                                                        <img class="prd_img">
                                                        <ul><p class="prd_brand">dailylike</p>
                                                        <p class="prd_name">�ų����� ����ī�� ��Ű��</p></ul>
                                                    </li>
                                                    <li>
                                                        <ul><p class="badge"></p>
                                                        <p class="info">���ǻ���</p></ul>
                                                    </li>
                                                    <li>
                                                        <ul>
                                                            <p class="percent">100%</p>
                                                            <ul class="bar">
                                                                <p class="bar_gray"></p>
                                                                <p class="bar_black"></p>
                                                            </ul>
                                                            <p class="">100/100</p>
                                                        </ul>
                                                    </li>
                                                   <li>
                                                        <ul>
                                                            <p class="info">�ٹ�����</p>
                                                            <p class="coupon"></p>
                                                        </ul>
                                                    </li>
                                                    <li>
                                                        <ul>
                                                            <button>����</button>
                                                            <button>����</button>
                                                        </ul>
                                                    </li>
                                                </ul>
                                            </div>
                                            <div class="tbody_wrap">
                                                <ul class="tbody">
                                                    <li><input type="checkbox"></li>
                                                    <li>12345678</li>
                                                    <li>
                                                        <img class="prd_img">
                                                        <ul><p class="prd_brand">dailylike</p>
                                                        <p class="prd_name">�ų����� ����ī�� ��Ű��</p></ul>
                                                    </li>
                                                    <li>
                                                        <ul><p class="badge"></p>
                                                        <p class="info">���ǻ���</p></ul>
                                                    </li>
                                                    <li>
                                                        <ul>
                                                            <p class="percent"></p>
                                                            <ul class="bar">
                                                                <p class="bar_gray"></p>
                                                                <p class="bar_black"></p>
                                                            </ul>
                                                            <p class="">50,000/50,000</p>
                                                        </ul>
                                                    </li>
                                                   <li>
                                                        <ul>
                                                            <p class="info"></p>
                                                            <p class="coupon">���ʽ����� ���� 3,000</p>
                                                        </ul>
                                                    </li>
                                                    <li>
                                                        <ul>
                                                            <button>����</button>
                                                            <button>����</button>
                                                        </ul>
                                                    </li>
                                                </ul>
                                            </div>
                                        </div>
                                    </div>
                                </div>
                            </div>
                        </div>
                        <div class="steps_bottom">
                            <ul>
                                <button class="delete">���</button>
                                <!-- ��ư ��Ȱ��ȭ : next on -->
                                <button class="next">����</button>
                            </ul>
                        </div>
                    </div>

                    <!-- ���̾��˾� ���� : paidGift_lyr on -->
                    <div class="lyr paidGift_lyr">
                        <div class="lyr_overlay"></div>
                     <!-- lyr_wrap lyr01 on -->
                        <!-- �׷� �߰��ϱ� -->
                        <div class="lyr_wrap lyr10 on">
                            <div class="lyr_top">
                                <li>�׷� �߰��ϱ�</li>
                            </div>
                            <div class="lyr_cont_wrap">
                                <div class="lyr_cont">
                                    <div class="cont_wrap search">
                                        <input type="text" class="input_l">
                                    </div>
                                </div>
                            </div>
                            <div class="lyr_bottom">
                                <ul>
                                    <button class="cancel">���</button>
                                    <button class="submit">Ȯ��</button>
                                </ul>
                            </div>
                        </div>
                        <!-- �׷� �����ϱ� -->
                        <div class="lyr_wrap lyr11">
                            <div class="lyr_top">
                                <li>�׷� �����ϱ�</li>
                            </div>
                            <div class="lyr_cont_wrap">
                                <div class="lyr_cont">
                                    <h4>�׷� �̸�</h4>
                                    <div class="cont_wrap search">
                                        <input type="text" class="input_l">
                                    </div>
                                </div>
                            </div>
                            <div class="lyr_bottom">
                                <button class="initial">�׷����</button>
                                <ul>
                                    <button class="cancel">���</button>
                                    <button class="submit">Ȯ��</button>
                                </ul>
                            </div>
                        </div>
                        <!-- �ݾ׺� �׷� �߰��ϱ� -->
                        <div class="lyr_wrap lyr12">
                            <div class="lyr_top">
                                <li>�ݾ׺� �׷� �߰��ϱ�</li>
                            </div>
                            <div class="lyr_cont_wrap">
                                <!-- ����� �̺�Ʈ - �ݾ׺� �׷� �߰��ϱ� -->
                                 <div class="lyr_cont" style="display:none;">
                                    <li class="type01 noti">���� ���� ������ ������̺�Ʈ�� ������ ���<br> �ݾ� ������ ������ ����� �̺�Ʈ�� ���ذ� �����ϰ� ����˴ϴ�.</li>
                                </div>
                                <div class="lyr_cont">
                                    <h4>�׷� �̸�<span class="gray">29/30</span></h4></h4>
                                    <div class="cont_wrap search">
                                        <input type="text" class="input_l">
                                    </div>
                                </div>
                                <div class="lyr_cont">
                                    <h4>�ݾ� ����</h4>
                                    <div class="cont_wrap">
                                        <input class="type04" type="checkbox" id="lyr12_01"><label for="lyr12_01">�� ���űݾ�</label>
                                        <input class="type04" type="checkbox" id="lyr12_02"><label for="lyr12_02">ī�װ�</label>
                                        <input class="type04" type="checkbox" id="lyr12_03"><label for="lyr12_03">�귣��</label>
                                        <input class="type04" type="checkbox" id="lyr12_04"><label for="lyr12_04">��ȹ��/�̺�Ʈ</label>
                                    </div>
                                </div>
                                <!-- �� ���űݾ� ���ý� ���� -->
                                <div class="lyr_cont">
                                    <h4>��� ����</h4>
                                    <input type="checkbox" id="lyr12_05" class="type01"><label for="lyr12_05"><span class="circle"></span>��ü��ǰ</label>
                                    <input type="checkbox" id="lyr12_06" class="type01"><label for="lyr12_06"><span class="circle"></span>�ٹ����� ��� ����</label>
                                </div>
                                <!-- ī�װ� ���ý� ���� -->
                                <div class="lyr_cont" style="display:block;">
                                    <h4>ī�װ� ����<span>*�������� �����մϴ�.</span></h4>
                                    <select name="" id="">
                                        <option value="">1depth</option>
                                    </select>
                                    <select name="" id="">
                                        <option value="">2depth</option>
                                    </select>
                                    <button class="add btn_blue">ī�װ� �߰�</button>
                                    <div class="option">
                                        <div class="option_added">
                                            <li>�����ι���</li>
                                            <li>></li>
                                            <li>���̾/�÷���</li>
                                            <li><button class="close"><img src="https://webadmin.10x10.co.kr/images/icon/close_gray.png" alt=""></button></li>
                                        </div>
                                        <div class="option_added">
                                            <li>�����ι���</li>
                                            <li>></li>
                                            <li>���ڷ��̼�</li>
                                            <li><button class="close"><img src="https://webadmin.10x10.co.kr/images/icon/close_gray.png" alt=""></button></li>
                                        </div>
                                    </div>
                                </div>
                                <!-- �귣�� ���ý� ���� -->
                                <div class="lyr_cont" style="display:none;">
                                    <h4>�귣�� ����<span>*�������� �����մϴ�.</span></h4>
                                    <div class="cont_wrap search brand">
                                         <input type="text" placeholder="�귣��ID�� �Է����ּ���">
                                        <button class="add btn_blue">�߰��ϱ�</button>
                                        <button class="search btn_white">�귣�� ã��</button>
                                    </div>
                                </div>
                                <!-- ��ȹ��/�̺�Ʈ ���ý� ���� -->
                                <div class="lyr_cont" style="display:none;">
                                    <h4>��ȹ��/�̺�Ʈ ����<span>*�������� �����մϴ�.</span></h4>
                                    <div class="cont_wrap search event">
                                        <input type="text" placeholder="�̺�Ʈ �ڵ带 �Է����ּ���">
                                        <button class="add btn_blue">�߰��ϱ�</button>
                                    </div>
                                </div>
                                <!-- ��ȹ��/�̺�Ʈ ���ý� ���� -->
                                 <div class="lyr_cont" style="display:block;">
                                    <h4>��ǰ ���� ����</h4>
                                    <input type="checkbox" id="lyr12_07" class="type01"><label for="lyr12_07"><span class="circle"></span>1���� ���� ��</label>
                                    <input type="checkbox" id="lyr12_08" class="type01"><label for="lyr12_08"><span class="circle"></span>��� ��ǰ ���� ��</label>
                                    <li class="type01 noti">������ ��ǰ �� 1���� �����ϸ� ������ �����մϴ�.</li>
                                </div>
                                <div class="lyr_cont">
                                    <h4>���űݾ� ����</h4>
                                    <li>�ֹ��ݾ���</li>
                                    <input type="text" id="lyr12_09" placeholder="0"><label for="lyr12_09"></label>
                                    <span>�̻��� ��� ���� ����</span>
                                </div>
                            </div>
                            <div class="lyr_bottom">
                                <ul>
                                    <button class="cancel">���</button>
                                    <button class="submit">Ȯ��</button>
                                </ul>
                            </div>
                        </div>
                        <!-- �ݾ׺� �׷� �����ϱ� -->
                        <div class="lyr_wrap lyr13">
                            <div class="lyr_top">
                                <li>�ݾ׺� �׷� �����ϱ�</li>
                            </div>
                            <div class="lyr_cont_wrap">
                                <!-- ����� �̺�Ʈ - �ݾ׺� �׷� �߰��ϱ� -->
                                 <div class="lyr_cont" style="display:none;">
                                    <li class="type01 noti">���� ���� ������ ������̺�Ʈ�� ������ ���<br> �ݾ� ������ ������ ����� �̺�Ʈ�� ���ذ� �����ϰ� ����˴ϴ�.</li>
                                </div>
                                <div class="lyr_cont">
                                    <h4>�׷� �̸�<span class="gray">29/30</span></h4></h4>
                                    <div class="cont_wrap search">
                                        <input type="text" class="input_l">
                                    </div>
                                </div>
                                <div class="lyr_cont">
                                    <h4>�ݾ� ����</h4>
                                    <div class="cont_wrap">
                                        <input class="type04" type="checkbox" id="lyr13_01"><label for="lyr13_01">�� ���űݾ�</label>
                                        <input class="type04" type="checkbox" id="lyr13_02"><label for="lyr13_02">ī�װ�</label>
                                        <input class="type04" type="checkbox" id="lyr13_03"><label for="lyr13_03">�귣��</label>
                                        <input class="type04" type="checkbox" id="lyr13_04"><label for="lyr13_04">��ȹ��/�̺�Ʈ</label>
                                    </div>
                                </div>
                                <!-- �� ���űݾ� ���ý� ���� -->
                                <div class="lyr_cont">
                                    <h4>��� ����</h4>
                                    <input type="checkbox" id="lyr13_05" class="type01"><label for="lyr13_05"><span class="circle"></span>��ü��ǰ</label>
                                    <input type="checkbox" id="lyr13_06" class="type01"><label for="lyr13_06"><span class="circle"></span>�ٹ����� ��� ����</label>
                                </div>
                                <!-- ī�װ� ���ý� ���� -->
                                <div class="lyr_cont" style="display:none;">
                                    <h4>ī�װ� ����<span>*�������� �����մϴ�.</span></h4>
                                    <select name="" id="">
                                        <option value="">1depth</option>
                                    </select>
                                    <select name="" id="">
                                        <option value="">2depth</option>
                                    </select>
                                    <button class="add btn_blue">ī�װ� �߰�</button>
                                    <div class="option">
                                        <div class="option_added">
                                            <li>�����ι���</li>
                                            <li>></li>
                                            <li>���̾/�÷���</li>
                                            <li><button class="close"><img src="https://webadmin.10x10.co.kr/images/icon/close_gray.png" alt=""></button></li>
                                        </div>
                                        <div class="option_added">
                                            <li>�����ι���</li>
                                            <li>></li>
                                            <li>���ڷ��̼�</li>
                                            <li><button class="close"><img src="https://webadmin.10x10.co.kr/images/icon/close_gray.png" alt=""></button></li>
                                        </div>
                                    </div>
                                </div>
                                <!-- �귣�� ���ý� ���� -->
                                <div class="lyr_cont" style="display:none;">
                                    <h4>�귣�� ����<span>*�������� �����մϴ�.</span></h4>
                                    <div class="cont_wrap search brand">
                                         <input type="text" placeholder="�귣��ID�� �Է����ּ���">
                                        <button class="add btn_blue">�߰��ϱ�</button>
                                        <button class="search btn_white">�귣�� ã��</button>
                                    </div>
                                </div>
                                <!-- ��ȹ��/�̺�Ʈ ���ý� ���� -->
                                <div class="lyr_cont" style="display:none;">
                                    <h4>��ȹ��/�̺�Ʈ ����<span>*�������� �����մϴ�.</span></h4>
                                    <div class="cont_wrap search event">
                                        <input type="text" placeholder="�̺�Ʈ �ڵ带 �Է����ּ���">
                                        <button class="add btn_blue">�߰��ϱ�</button>
                                    </div>
                                </div>
                                <!-- ��ȹ��/�̺�Ʈ ���ý� ���� -->
                                 <div class="lyr_cont" style="display:block;">
                                    <h4>��ǰ ���� ����</h4>
                                    <input type="checkbox" id="lyr13_07" class="type01"><label for="lyr13_07"><span class="circle"></span>1���� ���� ��</label>
                                    <input type="checkbox" id="lyr13_08" class="type01"><label for="lyr13_08"><span class="circle"></span>��� ��ǰ ���� ��</label>
                                    <li class="type01 noti">������ ��ǰ �� 1���� �����ϸ� ������ �����մϴ�.</li>
                                </div>
                                <div class="lyr_cont">
                                    <h4>���űݾ� ����</h4>
                                    <li>�ֹ��ݾ���</li>
                                    <input type="text" id="lyr13_09" placeholder="0"><label for="lyr13_09"></label>
                                    <span>�̻��� ��� ���� ����</span>
                                </div>
                            </div>
                            <div class="lyr_bottom">
                                <ul>
                                    <button class="cancel">���</button>
                                    <button class="submit">Ȯ��</button>
                                </ul>
                            </div>
                        </div>
                        <!-- �÷������� �����ϱ� -->
                        <div class="lyr_wrap lyr14">
                            <div class="lyr_top">
                                <li>�÷������� �����ϱ�</li>
                            </div>
                            <div class="lyr_cont_wrap">
                                <div class="lyr_cont">
                                    <h4>������ ��ǰ</h4>
                                    <div class="cont_wrap">
                                        <div class="table type02">
                                            <div class="tbody_wrap">
                                                <ul class="tbody">
                                                    <li>
                                                        <img class="prd_img">
                                                        <ul>
                                                        <p class="prd_code">12345678</p>
                                                        <p class="prd_brand">dailylike</p>
                                                        <p class="prd_name">�ų����� ����ī�� ��Ű��</p></ul>
                                                    </li>
                                                    <li>
                                                        <ul>
                                                            <p class="price p_type01"><span>38,000</span></p>
                                                            <p class="price p_type02">7% ����<span>38,000</span></p>
                                                            <p class="price p_type03">7% ����<span>38,000</span></p>
                                                        </ul>
                                                    </li>
                                                    <li>
                                                        <ul>
                                                            <p class="p_type01">38,000</p>
                                                            <p class="p_type02">23,902</p>
                                                            <p class="p_type03">23,902</p>
                                                        </ul>
                                                    </li>
                                                    <li>
                                                        <ul>
                                                            <p class="p_type01">27%</p>
                                                            <p class="p_type02">26%</p>
                                                            <p class="p_type03">26%</p>
                                                        </ul>
                                                    </li>
                                                </ul>
                                            </div>
                                        </div>
                                    </div>
                                </div>
                                 <div class="lyr_cont">
                                    <h4>�÷��� ���ΰ� ����</h4>
                                    <div class="option t02">
                                        <ul>
                                            <li>���ΰ�</li>
                                            <input type="text" id="lyr14_01"><label for="lyr14_01"></label>
                                        </ul>
                                        <ul class="percent">
                                            <li>������</li>
                                            <input type="text" id="lyr14_02"><label for="lyr14_02"></label>
                                        </ul>
                                    </div>
                                </div>
                                <div class="lyr_cont">
                                    <h4>���� �� ������</h4>
                                    <select name="" id="">
                                        <option value="">�ٹ����ٺδ�</option>
                                        <option value="">�ٹ����ٺδ�</option>
                                    </select>
                                    <span>(�ǸŰ�) 32,300</span>
                                    <span>(���θ��԰�) 23,902</span>
                                    <b>26%</b>
                                </div>
                                <div class="lyr_cont">
                                    <h4>����ǰ ���� ����</h4>
                                    <li class="type02"><input class="type02" type="checkbox" id="lyr14_03" class="type01" ><label for="lyr14_03"><span class="circle"></span>���� ����<input type="text" placeholder="0">��</li>
                                    <li class="type02"><input class="type02" type="checkbox" id="lyr14_04" class="type01"><label for="lyr14_04"><span class="circle"></span>������</label></li>
                                </div>
                                <div class="lyr_cont">
                                    <h4>�ִ� ���� �� ����</h4>
                                    <div class="option t03 lyr14_05">
                                        <input type="text" id="lyr14_05"><label for="lyr14_05"></label>
                                    </div>
                                </div>
                                <div class="lyr_cont">
                                    <button class="badge_add">
                                        <li>Ư���� ��ǰ�� ���� �ѳ��� ����<br>���� �� ���ǻ��� �߰��ϱ�</li>
                                        <li class="arrow"></li>
                                    </button>
                                </div>
                                <!-- badge_add ���� �߰� ��ư Ŭ���� -->
                                <div class="lyr_cont badge">
                                    <div class="cont_wrap">
                                        <ul>
                                            <h4>��ǰ ���� ����<span class="gray">29/30</span></h4>
                                            <input type="text" placeholder="������ �� ������ �Է����ּ���">
                                            <h4 class="badge_noti">���ǻ���<span class="gray">29/30</span></h4>
                                            <textarea placeholder="��ǰ ���� �ϴܿ� ǥ�õ� ���ǻ������Է����ּ���"></textarea>
                                        </ul>
                                        <ul>
                                            <h4>������ ���ǻ��� �̸�����</h4>
                                            <li class="badge_img"><img src="https://webadmin.10x10.co.kr/images/icon/unit.png"></li>
                                        </ul>
                                    </div>
                                </div>
                            </div>
                            <div class="lyr_bottom">
                                <ul>
                                    <button class="cancel">���</button>
                                    <button class="submit">Ȯ��</button>
                                </ul>
                            </div>
                        </div>
                        <!-- ����ǰ �űԵ�� -->
                        <div class="lyr_wrap lyr15">
                            <div class="lyr_top">
                                <li>����ǰ �űԵ��</li>
                            </div>
                            <div class="lyr_cont_wrap">
                                <div class="lyr_cont">
                                    <h4>����ǰ ����</h4>
                                    <div class="cont_wrap type04">
                                        <input class="type04" type="checkbox" id="lyr15_01"><label for="lyr15_01">��ǰ</label>
                                        <input class="type04" type="checkbox" id="lyr15_02"><label for="lyr15_02">���ʽ� ����</label>
                                        <input class="type04" type="checkbox" id="lyr15_03"><label for="lyr15_03">���ϸ���</label>
                                    </div>
                                    <li class="noti">���ϸ����� ��ü ���� �̺�Ʈ������ ����� �� �ֽ��ϴ�.</li>
                                </div>
                                <div class="lyr_cont">
                                    <h4>����ǰ��</h4>
                                    <div class="cont_wrap search">
                                        <input type="text" class="input_l" placeholder="����ǰ���� �Է����ּ���">
                                    </div>
                                </div>
                                <!-- ��ǰ ���ý� ���� -->
                                <div class="lyr15_1" style="display:none;">
                                    <div class="lyr_cont">
                                        <h4>��۹��</h4>
                                        <input type="checkbox" id="lyr15_04" class="type01"><label for="lyr15_04"><span class="circle"></span>�ٹ����� ���</label>
                                        <input type="checkbox" id="lyr15_05" class="type01"><label for="lyr15_05"><span class="circle"></span>��ü���</label>
                                    </div>
                                    <div class="lyr_cont">
                                        <h4>�귣�� ID</h4>
                                        <div class="cont_wrap search brand">
                                            <input type="text" placeholder="�귣��ID�� �Է����ּ���">
                                            <button class="add btn_blue">�귣�� ID �˻�</button>
                                        </div>
                                    </div>
                                    <div class="lyr_cont">
                                        <h4>��÷��ǰ�ڵ�</h4>
                                        <div class="cont_wrap search">
                                            <input type="text" class="input_l" placeholder="����ǰ���� �Է����ּ���">
                                        </div>
                                    </div>
                                    <div class="lyr_cont">
                                        <h4>�����ڵ�</h4>
                                        <div class="cont_wrap search code">
                                            <input type="text" class="code01">
                                            <input type="text" class="code02">
                                            <input type="text" class="code03">
                                            <button class="add btn_blue btn01">�˻�</button>
                                            <button class="add btn_blue btn02">����ǰ �����ڵ� �ڵ�����</button>
                                        </div>
                                    </div>
                                    <div class="lyr_cont">
                                        <h4>��ǥ �̹���</h4>
                                        <li class="noti">����ǰ ������ ���� �̹����� Ȱ��Ǵ� ����� �����մϴ�.</li>
                                        <div class="cont_wrap img">
                                            <button class="add"><li class="plus"><img src="https://webadmin.10x10.co.kr/images/icon/plus.png"></li>�̹��� ���</button>
                                        </div>
                                    </div>
                                </div>
                                <!-- ���ʽ� ���� ���ý� ���� -->
                                <div class="lyr15_2" style="display:none;">
                                    <div class="lyr_cont">
                                        <h4>�����ڵ�</h4>
                                        <div class="cont_wrap search">
                                            <input type="text" class="input_l"placeholder="������ ���� �ڵ带 �Է����ּ���">
                                        </div>
                                    </div>
                                    <div class="lyr_cont">
                                        <h4>��ǥ �̹���</h4>
                                        <li class="noti">����ǰ ������ ���� �̹����� Ȱ��˴ϴ�.</li>
                                        <li class="type02">
                                            <input class="type02" type="checkbox" id="lyr15_06" ><label for="lyr15_06"><span class="circle"></span>�⺻ �̹���</label>
                                            <div class="cont_wrap img">
                                                <img src="https://webadmin.10x10.co.kr/images/icon/coupon.png">
                                            </div>
                                        </li>
                                       <li class="type02">
                                            <input class="type02" type="checkbox" id="lyr15_07" ><label for="lyr15_07"><span class="circle"></span>�̹��� ���� ���</label>
                                            <div class="cont_wrap img">
                                                <button class="add"><li class="plus"><img src="https://webadmin.10x10.co.kr/images/icon/plus.png"></li>�̹��� ���</button>
                                            </div>
                                       </li>
                                    </div>
                                </div>
                                <!-- ���ϸ��� ���ý� ���� -->
                                <div class="lyr15_3" style="display:block;">
                                    <div class="lyr_cont">
                                        <h4>���ϸ��� ���� �ݾ�</h4>
                                        <div class="cont_wrap search">
                                            <input type="text" class="input_l"placeholder="�ݾ��� �Է����ּ���">
                                        </div>
                                    </div>
                                    <div class="lyr_cont">
                                        <h4>��ȿ�Ⱓ ����</h4>
                                        <ul>
                                            <li class="type02">
                                            <input class="type02" type="checkbox" id="lyr15_08" ><label for="lyr15_08"><span class="circle"></span>�Ⱓ����</label>
                                            </li>
                                            <li class="type02">
                                                <input class="type02" type="checkbox" id="lyr15_09" ><label for="lyr15_09"><span class="circle"></span>ȸ������ ����</label>
                                            </li>
                                        </ul>
                                        <!-- �Ⱓ���� -->
                                        <div class="option lyr15_08">
                                            <input type="text" placeholder="30"> �� ���� ���ϸ��� �Ҹ� 
                                        </div>
                                        <!-- ȸ������ ���� -->
                                        <div class="option">
                                            <div class="date_wrap">
                                                <ul class="date">
                                                    <button class="date_btn"><input type="text" placeholder="2022.01.01" class="" readonly><img src="https://webadmin.10x10.co.kr/images/icon/calendar.png" alt="">
                                                    <!-- ��¥ ���� -->
                                                    <div class="cal_month t02" style="display:none;">
                                                            <div class="arrow"></div>
                                                                <table class="table-condensed table-bordered table-striped">
                                                                         <thead>
                                                                            <tr>
                                                                                <th colspan="7">
                                                                                    <ul class="btn_group">
                                                                                        <li class="btn"><img class="arrow_gray left" src="https://webadmin.10x10.co.kr/images/icon/arrow_gray.png"></li>
                                                                                        <li class="btn active">2�� 2022</li>
                                                                                        <li class="btn"><img class="arrow_gray" src="https://webadmin.10x10.co.kr/images/icon/arrow_gray.png"></li>
                                                                                    </ul>
                                                                                </th>
                                                                            </tr>
                                                                        </thead>
                                                                        <tbody>
                                                                            <tr>
                                                                                <td class="gray">30</td>
                                                                                <td class="gray">31</td>
                                                                                <td class="on">1</td>
                                                                                <td>2</td>
                                                                                <td>3</td>
                                                                                <td>4</td>
                                                                                <td>5</td>
                                                                            </tr>
                                                                            <tr>
                                                                                <td>6</td>
                                                                                <td>7</td>
                                                                                <td>8</td>
                                                                                <td>9</td>
                                                                                <td>10</td>
                                                                                <td>11</td>
                                                                                <td>12</td>
                                                                            </tr>
                                                                            <tr>
                                                                                <td>13</td>
                                                                                <td>14</td>
                                                                                <td>15</td>
                                                                                <td>16</td>
                                                                                <td>17</td>
                                                                                <td>18</td>
                                                                                <td>19</td>
                                                                            </tr>
                                                                            <tr>
                                                                                <td>20</td>
                                                                                <td>21</td>
                                                                                <td>22</td>
                                                                                <td>23</td>
                                                                                <td>24</td>
                                                                                <td>25</td>
                                                                                <td>26</td>
                                                                            </tr>
                                                                            <tr>
                                                                                <td>27</td>
                                                                                <td>28</td>
                                                                                <td class="gray">1</td>
                                                                                <td class="gray">2</td>
                                                                                <td class="gray">3</td>
                                                                                <td class="gray">4</td>
                                                                                <td class="gray">5</td>
                                                                            </tr>
                                                                        </tbody>
                                                            </table>
                                                    </div>
                                                    <input type="text" placeholder="00" class="time">:<input type="text" placeholder="00" class="time"><img src="https://webadmin.10x10.co.kr/images/icon/clock.png">
                                                </ul>
                                            </div>
                                        </div>
                                    </div>
                                    <div class="lyr_cont">
                                        <h4>��ǥ �̹���</h4>
                                        <li class="noti">����ǰ ������ ���� �̹����� Ȱ��˴ϴ�.</li>
                                        <ul>
                                            <li class="type02">
                                                <input class="type02" type="checkbox" id="lyr15_10" ><label for="lyr15_10"><span class="circle"></span>�⺻ �̹���</label>
                                                <div class="cont_wrap img">
                                                    <img src="https://webadmin.10x10.co.kr/images/icon/mileage.png">
                                                </div>
                                            </li>
                                            <li class="type02">
                                                <input class="type02" type="checkbox" id="lyr15_11" ><label for="lyr15_11"><span class="circle"></span>�̹��� ���� ���</label>
                                                <div class="cont_wrap img">
                                                    <button class="add"><li class="plus"><img src="https://webadmin.10x10.co.kr/images/icon/plus.png"></li>�̹��� ���</button>
                                                </div>
                                            </li>
                                       </ul>
                                    </div>
                                </div>
                            </div>
                            <div class="lyr_bottom">
                                <ul>
                                    <button class="cancel">���</button>
                                    <button class="submit">Ȯ��</button>
                                </ul>
                            </div>
                        </div>
                        <!-- ����ǰ ���� -->
                        <div class="lyr_wrap lyr16">
                            <div class="lyr_top">
                                <li>����ǰ ����</li>
                            </div>
                            <!-- ��ǰ -->
                            <div class="lyr_cont_wrap" style="display:none;">
                                <div class="lyr_cont">
                                    <h4>����ǰ ����</h4>
                                    <div class="add_btn">
                                        <button>����ǰ �űԵ��</button><span></span>
                                        <button>���� ����ǰ �ҷ�����</button>
                                    </div>
                                    <div class="cont_wrap">
                                        <div class="table type02">
                                        <a>
                                            <div class="tbody_wrap">
                                                <ul class="tbody">
                                                    <li>
                                                        <ul>
                                                            <span>��ǰ</span>
                                                            <p>12345678</p>
                                                        </ul>
                                                    </li>
                                                    <li>
                                                        <img class="prd_img">
                                                        <ul>
                                                        <p class="prd_brand">dailylife</p>
                                                        <p class="prd_name">�ų����� ����ī�� ��Ű����Ű����Ű��</p></ul>
                                                    </li>
                                                    <li>
                                                        <ul>
                                                            <p>��ü��۾�ü��۾�ü���</p>
                                                        </ul>
                                                    </li>
                                                </ul>
                                                </a>
                                            </div>
                                        </div>
                                    </div>
                                </div>
                                <div class="lyr_cont">
                                    <h4>���� Ƚ�� ����</h4>
                                    <li class="type02"><input class="type02" type="checkbox" id="lyr16_01" ><label for="lyr16_01"><span class="circle"></span>�������� <input type="text" placeholder="0" > ��</label></li>
                                    <li class="type02"><input class="type02" type="checkbox" id="lyr16_02" ><label for="lyr16_02"><span class="circle"></span>������</label></li>
                                </div>
                                <div class="lyr_cont">
                                    <h4>����ǰ �̹��� ����</h4>
                                    <div class="cont_wrap img">
                                        <li>���� ������ �����</li>
                                        <ul>
                                            <button class="add"><li class="plus"><img src="https://webadmin.10x10.co.kr/images/icon/plus.png"></li>�̹��� ���</button>
                                        </ul>
                                        <li>�� �˾� �̹���</li>
                                        <ul class="img_wrap">
                                            <button class="add"><li class="plus"><img src="https://webadmin.10x10.co.kr/images/icon/plus.png"></li>�̹��� ���</button>
                                            <button class="add"><li class="plus"><img src="https://webadmin.10x10.co.kr/images/icon/plus.png"></li>�̹��� ���</button>
                                            <button class="add"><li class="plus"><img src="https://webadmin.10x10.co.kr/images/icon/plus.png"></li>�̹��� ���</button>
                                            <button class="add"><li class="plus"><img src="https://webadmin.10x10.co.kr/images/icon/plus.png"></li>�̹��� ���</button>
                                            <button class="add"><li class="plus"><img src="https://webadmin.10x10.co.kr/images/icon/plus.png"></li>�̹��� ���</button>
                                        </ul>
                                    </div>
                                </div>
                            </div>
                            <!-- ���ʽ����� -->
                            <div class="lyr_cont_wrap" style="display:none;">
                                <div class="lyr_cont">
                                    <h4>����ǰ ����</h4>
                                    <div class="add_btn">
                                            <button>����ǰ �űԵ��</button><span></span>
                                            <button>���� ����ǰ �ҷ�����</button>
                                    </div>
                                    <div class="cont_wrap">
                                        <div class="table type02">
                                        <a>
                                            <div class="tbody_wrap">
                                                <ul class="tbody">
                                                    <li>
                                                        <ul>
                                                            <span>���ʽ�����</span>
                                                            <p>12345678</p>
                                                        </ul>
                                                    </li>
                                                    <li>
                                                        <img class="prd_img">
                                                        <ul>
                                                        <p>10x10giftmall</p>
                                                        <p class="prd_name">�ٹ������� ���! ��ۺ� ���� ����!</p></ul>
                                                    </li>
                                                    <li>
                                                        <ul>
                                                            <p>���ʽ�����<br>3000�� ����</p>
                                                        </ul>
                                                    </li>
                                                </ul>
                                                </a>
                                            </div>
                                        </div>
                                    </div>
                                </div>
                                <div class="lyr_cont">
                                    <h4>���� Ƚ�� ����</h4>
                                    <li class="type02"><input class="type02" type="checkbox" id="lyr16_03" ><label for="lyr16_03"><span class="circle"></span>�������� <input type="text" placeholder="0" > ��</label></li>
                                    <li class="type02"><input class="type02" type="checkbox" id="lyr16_04" ><label for="lyr16_04"><span class="circle"></span>������</label></li>
                                </div>
                                <div class="lyr_cont">
                                    <h4>����ǰ �̹��� ����</h4>
                                    <div class="cont_wrap img">
                                        <li>���� ������ �����</li>
                                        <ul>
                                            <button class="add"><li class="plus"><img src="https://webadmin.10x10.co.kr/images/icon/plus.png"></li>�̹��� ���</button>
                                        </ul>
                                    </div>
                                </div>
                            </div>
                            <!-- ���ϸ��� -->
                            <div class="lyr_cont_wrap" style="display:block;">
                                <div class="lyr_cont">
                                    <h4>����ǰ ����</h4>
                                    <div class="add_btn">
                                            <button>����ǰ �űԵ��</button><span></span>
                                            <button>���� ����ǰ �ҷ�����</button>
                                    </div>
                                    <div class="cont_wrap">
                                        <div class="table type02">
                                        <a>
                                            <div class="tbody_wrap">
                                                <ul class="tbody">
                                                    <li>
                                                        <ul>
                                                            <span>���ϸ���</span>
                                                            <p>12345678</p>
                                                        </ul>
                                                    </li>
                                                    <li>
                                                        <img class="prd_img">
                                                        <ul>
                                                        <p>10x10giftmall</p>
                                                        <p class="prd_name">�ٹ������� ���! 1,010������!</p></ul>
                                                    </li>
                                                    <li>
                                                        <ul>
                                                            <p>���ϸ���<br>1,010P ����</p>
                                                        </ul>
                                                    </li>
                                                </ul>
                                                </a>
                                            </div>
                                        </div>
                                    </div>
                                </div>
                                <div class="lyr_cont">
                                    <h4>���� Ƚ�� ����</h4>
                                    <li class="type02"><input class="type02" type="checkbox" id="lyr16_05" ><label for="lyr16_05"><span class="circle"></span>�������� <input type="text" placeholder="0" > ��</label></li>
                                    <li class="type02"><input class="type02" type="checkbox" id="lyr16_06" ><label for="lyr16_06"><span class="circle"></span>������</label></li>
                                </div>
                                <div class="lyr_cont">
                                    <h4>����ǰ �̹��� ����</h4>
                                    <div class="cont_wrap img">
                                        <li>���� ������ �����</li>
                                        <ul>
                                            <button class="add"><li class="plus"><img src="https://webadmin.10x10.co.kr/images/icon/plus.png"></li>�̹��� ���</button>
                                        </ul>
                                    </div>
                                </div>
                            </div>
                            <div class="lyr_bottom">
                                <ul>
                                    <button class="cancel">���</button>
                                    <button class="submit">Ȯ��</button>
                                </ul>
                            </div>
                        </div>
                    </div>
                </div>
                <script>
                </script>
            </div>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
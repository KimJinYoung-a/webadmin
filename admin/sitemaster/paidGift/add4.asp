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
                            <li class="step on"><span></span>��������</li>
                        </div>
                        <div class="step_wrap step04">
                            <div class="step_noti on"><span>�����ϱ� �� ������ Ȯ���� ���ּ���!</span></div>
                            <div class="bg">
                                <div class="step_cont step04_01 on">
                                    <h2>�⺻���� �Է�</h2>
                                    <div class="step_detail">
                                        <h3>����ä�� ����</h3>
                                        <div class="step_detail_list">
                                            <input type="checkbox" class="type01" id="step04_01"><label for="step04_01"><span class="circle"></span>PC WEB</label>
                                            <input type="checkbox" class="type01" id="step04_02"><label for="step04_02"><span class="circle"></span>App(iOS/Android)</label>
                                            <input type="checkbox" class="type01" id="step04_03"><label for="step04_03"><span class="circle"></span>Mobile WEB</label>
                                        </div>
                                    </div>
                                    <div class="step_detail">
                                        <h3>����<span>22/99</span></h3>
                                        <textarea>�������ð����� ��Ī ��� �������� �׽�Ʈ</textarea>
                                    </div>
                                    <div class="step_detail">
                                        <h3>������<span>22/99</span></h3>
                                        <textarea>�������ð����� ��Ī ��� �������� �׽�Ʈ</textarea>
                                    </div>
                                    <div class="step_detail">
                                        <h3>�����</h3>
                                        <div class="step_detail_list">
                                            <li>
                                                <input type="text" value="�������� - ������">
                                                <a href="" class="close"><img class="arrow_gray left" src="https://webadmin.10x10.co.kr/images/icon/close_gray.png"></a>
                                            </li>
                                            <button class="btn_blue">����</button>
                                        </div>
                                    </div>
                                    <div class="step_detail">
                                        <h3>��뿩��</h3>
                                        <div class="step_detail_list">
                                            <input type="checkbox" class="type01" id="step04_04"><label for="step04_04"><span class="circle"></span>�����</label>
                                            <input type="checkbox" class="type01" id="step04_05"><label for="step04_05"><span class="circle"></span>������</label>
                                        </div>
                                    </div>
                                </div>
                                <div class="step_cont step04_02 on">
                                    <h2>�ֹ�/���� ȭ�� Ȯ��<li><span>2022.01.03 - 2022.01.20</span> ����Ǵ� ���������� ������ �������ּ���</li></h2>
                                    <div class="step_detail">
                                        <div class="box">
                                            <a href=""><img src="https://webadmin.10x10.co.kr/images/icon/drag.png"></a>
                                            <ul>
                                                <li>�긮�� ��Ī��� ���</li>
                                                <li class="period">2022.01.13 - 2022.01.20</li>
                                                <li class="info">���űݾ� ����</li>
                                            <ul>
                                        </div>
                                        <div class="box on">
                                            <a href=""><img src="https://webadmin.10x10.co.kr/images/icon/drag.png"></a>
                                            <ul>
                                                <li>�������ð����� ��Ī ��� �������� �׽�Ʈ</li>
                                                <li class="period">2022.01.13 - 2022.01.20</li>
                                                <li class="info">���űݾ� ����</li>
                                            <ul>
                                        </div>
                                        <div class="box">
                                            <a href=""><img src="https://webadmin.10x10.co.kr/images/icon/drag.png"></a>
                                            <ul>
                                                <li>�긮�� ��Ī��� ���</li>
                                                <li class="period">2022.01.13 - 2022.01.20</li>
                                                <li class="info">���űݾ� ����</li>
                                            <ul>
                                        </div>
                                        <div class="box">
                                            <a href=""><img src="https://webadmin.10x10.co.kr/images/icon/drag.png"></a>
                                            <ul>
                                                <li>�긮�� ��Ī��� ���</li>
                                                <li class="period">2022.01.13 - 2022.01.20</li>
                                                <li class="info">���űݾ� ����</li>
                                            <ul>
                                        </div>
                                        <div class="box">
                                            <a href=""><img src="https://webadmin.10x10.co.kr/images/icon/drag.png"></a>
                                            <ul>
                                                <li>�긮�� ��Ī��� ���</li>
                                                <li class="period">2022.01.13 - 2022.01.20</li>
                                                <li class="info">���űݾ� ����</li>
                                            <ul>
                                        </div>
                                    </div>
                                </div>
                            </div>
                            <div class="step_cont step04_03 on">
                                <h2>�ȳ�����<li>�ڼ��� �ȳ��ϰ��� �ϴ� ������ �ִٸ� �ۼ����ּ���. �˾����� ����˴ϴ�!</li></h2>
                                <div class="step_detail">
                                    <textarea>�� 2022 ���̾ ���丮 ��ǰ ���� 20,000�� �̻� ���� �� �����˴ϴ�.&#10;�� ����, ����ī�� �� ��� �� ����Ȯ�� �ݾ� �����Դϴ�.&#10;�� 2022 ���̾ ���丮 ��ǰ ���� 20,000�� �̻� ���� �� �����˴ϴ�.&#10;�� ����, ����ī�� �� ��� �� ����Ȯ�� �ݾ� �����Դϴ�.&#10;�� 2022 ���̾ ���丮 ��ǰ ���� 20,000�� �̻� ���� �� �����˴ϴ�.&#10;�� ����, ����ī�� �� ��� �� ����Ȯ�� �ݾ� �����Դϴ�.&#10;
                                    </textarea>
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
                </div>
                <script>
                </script>
            </div>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
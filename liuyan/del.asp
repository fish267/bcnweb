<!--
---------------------��    ��:del.asp
---------------------��������:���԰����(��ȫע��)
---------------------��������:Dmlt
---------------------���߲���:Http://HI.Baidu.com/Int_yi
---------------------���߸���:�����BUG,����©д����,����������ϵ;
---------------------��ϵ��ʽ:dmlt@vip.qq.com
-->

<!--#include file="conn.asp" -->
<!--#include file="sqlin.asp" -->
<%
    if session("pass")="" then
	response.write "��û�е�½��֤��!"
	response.end
    end if
%>

<%
id=request("id")                                        ' ��ȡID
conn.execute("delete from data where id in ("&id&")")   'ɾ����ѯ����ID
response.write "<script>history.go(-1)</script>"        '���ص���ҳ
%>
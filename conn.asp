<%
Dim Conn,ConnStr 
dim Sql_Server,Sql_User,Sql_Pass,Sql_Data 
Sql_Server = "127.0.0.1" '���ݿ��������ַ 
Sql_User = "sa" '���ݿ��¼�� 
Sql_Pass = "521521" '���ݿ����� 
Sql_Data = "uu00_dj" '���ݿ���
connstr = "PROVIDER=SQLOLEDB;DATA SOURCE="&Sql_Server&";UID="&Sql_User&";PWD="&Sql_Pass&";DATABASE="&Sql_Data 
set conn = server.createobject("ADODB.connection") '�������ݿ����Ӷ��� 
conn.open connstr '�������ݿ�
%>


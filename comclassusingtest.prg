* �������� ������������� ���������� COM-������
SET OLEOBJECT ON

obj = CREATEOBJECT("MyComLib.MyComClass")
result = obj.DoSomeWork("Dude")

MESSAGEBOX("The result: " + result)
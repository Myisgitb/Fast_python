'''

Private Sub Command1_Click()
    Set PythonUtils = CreateObject("PythonDemos.Utilities")
    Dim dict
    Set dict = CreateObject("Scripting.Dictionary")


    response = PythonUtils.SplitString("Hello from VB")
    ls = PythonUtils.List1("Hello from VB")
    Set dict = PythonUtils.Dict1("Hello from VB")
    sc = PythonUtils.Strcha("Hello from VB")
    ic = PythonUtils.Intcha("Hello from VB")

    For Each Item In response
        Form1.Print "strcha:"; Item
    Next
    For Each Item In ls
        Form1.Print "List:"; Item
    Next
    Form1.Print sc
    Form1.Print ic
    
    k = dict.keys
    v = dict.Items
    For Item = 0 To dict.Count - 1
        Form1.Print k(Item) & "：" & v(Item)
        
    Next


End Sub

Private Sub Command2_Click()
Unload Me
End Sub
'''
class PythonUtilities:
    _public_methods_=['Strcha','List1','Dict1','Intcha','SplitString']
    _reg_progid_='PythonDemos.Utilities'
    # 使用"print (pythoncom.CreateGuid())" 得到一个自己的clsid，不要用下面这个！！
    _reg_clsid_='{d23b6bda-3f51-4671-8ba9-becc4a7fd530}'
    def SplitString(self, val, item=None):
        import string 
        if item !=None: 
            item=str(item)
        val=str(val)
        return val.split(item)
    def Strcha(self, val, item=None):
        sc="this is Character string" 
        return sc
    def Intcha(self, val, item=None):
        Ic=3.1415926 
        return Ic
    def List1(self, val, item=None):
         la=["List",1979]
         return la
    def Dict1(self, val, item=None):
        import win32com.client
        Da = win32com.client.Dispatch("Scripting.Dictionary")
        Da["d01"]=2016
        Da["d02"]=2017
        Da["d03"]=2018
        Da["d04"]=2019
        Da["d05"]=2020
        return Da
if __name__=='__main__':
    print ('Registering COM server...')
    import win32com.server.register
    win32com.server.register.UseCommandLine(PythonUtilities)

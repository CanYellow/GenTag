;BrightHuang
;24.03.11
;根据csv文件在文件中插入特定格式的内容
;该文件需要用 utf-8 带BOM签名的格式保存，正常的utf-8格式中文会乱码

#Warn All, Off

FileName :=""
InputXH := 134481924
InputEng := 67699721

;按下Ctrl + shift \键，发送\原义字符
^+\::SendInput("\")

;按下\键，触发选择框
:*?:\::
{

;获取主窗口属性的准备工作
global FileName
global InputXH
global InputEng

uid := WinActive("A")
MouseGetPos(&xpos ,&ypos)

;获取输入法ID
InID := DllCall("GetKeyboardLayout","UINT",DllCall("GetWindowThreadProcessId","UINT", WinActive("A"),"UINTP",0))
;判断并保存输入法状态
Finput := (InID = InputXH)*2 + (InID = InputEng)*1

MyMap := Map()
ArrayIn :=[]

;未引入csv时，先手动引入csv文件
if(FileName="")
{
  FileName := FileSelect(3,"D:\Work\Workdirectory", , "*.csv")
}

;每次打开窗口前先载入数据
Read_Data(FileName)

;创建GUI
MyGui:=Gui(,"Auto match tag")
MyEdit:=MyGui.Add("Edit","-WantReturn")
MyLBox:= MyGui.Add("ListBox", "r15", ArrayIn)
CsvBtn:=MyGui.Add("Button",,"csv file")
ExitBtn:=MyGui.Add("Button","Default x+25","Exit")
pos := "X" . xpos . " " . "Y" . ypos
MyGui.Show(pos)

;GUI事件绑定
ExitBtn.OnEvent("Click", Gui_Close)
CsvBtn.OnEvent("Click", Load_File)
MyEdit.OnEvent("Change",Update_LBox)

Gui_Close(*){
  MyGui.Destroy()
}

Load_File(*){

  FileName := FileSelect(3,"D:\Work\Workdirectory", , "*.csv")
  ArrayIn :=[]
  MyMap.Clear()
  Read_Data(FileName)
  MyLBox.Delete()
  MyLBox.Add(ArrayIn)
}

Read_Data(CsvFile){
  if CsvFile {
    Loop Read, CsvFile
    {
      LineNumber := A_Index
      Loop parse, A_LoopReadLine, "CSV"
      {
        if(A_Index = 1)
          cid := A_LoopField
        if(A_Index = 2)
          name := A_LoopField
      }
      ArrayIn.Push(cid . ": " . name)
      MyMap[cid] := name
    }
  }
}

Update_LBox(*){
  MyStr:=MyEdit.Value
  ;msgbox(Instr)
  
  ;以下是根据输入动态刷新list的逻辑
  ;默认以输入的最后一位确定上屏格式，这意味着两种情况
  ;若只有一位，直接匹配刷新
  ;若超出一位，则截尾刷新
  
  ;可选的选项集合
  OptList := "abcnv"
  Len := StrLen(MyStr)
  ;msgbox(MyStr . " " . len)
  if(Len <= 1) {
    PreCode := MyStr
    SubOpt := ""
  }
  else if (Len > 1){
    ;末尾的格式选项符号
    SubOpt := SubStr(MyStr, -1)
    if (Not InStr(OptList, SubOpt)){
      PreCode := MyStr
      SubOpt := ""
    } else{
      ;序号查找所用的子串
      PreCode := SubStr(MyStr, 1, -1)
    }
  }
  
  ;更新List
  CanArr:=[]
  for idx, name in MyMap {
    if(PreCode != ""){
      res := InStr(idx, PreCode)
      If res {
        CanArr.Push(idx . ": " . name)
      }
    }
  }
  MyLBox.Delete()
  MyLBox.Add(CanArr)
  
  
  ;根据格式选项上屏
  if(SubOpt != ""){
      MyGui.Destroy()
      ;这是为了保证按键可靠发送到原始的编辑窗口，而不发送到AHK生成的窗口
      sleep 75
      SendStr(MyMap , PreCode, SubOpt, Finput)
  }
  
  SendStr(MyMap , Cod, Opt, Flg){
    Switch Opt {
      case "a":
        SendAll(MyMap, Flg)
      case "v":
        SendV(MyMap, Cod, Opt)
      case "b":
        SendB(MyMap, Cod, Opt)
      case "n":
        SendN(MyMap, Cod, Opt)
      case "c":
        SendC(MyMap, Cod, Opt)
      Default:
    }
  }

  ;a选项下，通过剪贴板全部上屏
  SendAll(MyMap, Flg){
    MySendStr :=""
    for idx, name in MyMap {
      res := InStr(idx, "#")
      If (Not res) {
        MySendStr .= idx . "：" . name . "；"
      }
    }
;使用SendInput，在WPS内是可行的，但是在Word里面是按字符逐个上屏，不能控制时间，也就不能准确控制输入法切换与输入之间的先后顺序
    ;索性使用ClipBoard工具
    A_Clipboard := SubStr(MySendStr, 1, -1) . "。"
    SendInput("{Ctrl down}v{Ctrl up}")
  }
  
  SendV(MyMap, Cod, Opt){
    MySendStr :=""
    for idx, name in MyMap {
      ;res := InStr(idx, Cod)
      ;If (res) {
      ;  MySendStr .= idx
      ;  break
      ;}
      If (idx = Cod) {
        MySendStr .= idx
        break
      }
    }
    A_Clipboard := MySendStr
    SendInput("{Ctrl down}v{Ctrl up}")
  }

  SendB(MyMap, Cod, Opt){
    MySendStr :=""
    for idx, name in MyMap {
      ;res := InStr(idx, Cod)
      ;If (res) {
      ;  MySendStr .= name . idx
      ;  break
      ;}
      If (idx = Cod) {
        MySendStr .= name . idx
        break
      }
    }
    A_Clipboard := MySendStr
    SendInput("{Ctrl down}v{Ctrl up}")
  }
  
  SendN(MyMap, Cod, Opt){
    MySendStr :=""
    for idx, name in MyMap {
      ;res := InStr(idx, Cod)
      ;If (res) {
      ;  MySendStr .= name
      ;  break
      ;}
      If (idx = Cod) {
        MySendStr .= name
        break
      }
    }
    A_Clipboard := MySendStr
    SendInput("{Ctrl down}v{Ctrl up}")
  }
  
  SendC(MyMap, Cod, Opt){
    MySendStr :=""
    for idx, name in MyMap {
      ;res := InStr(idx, Cod)
      ;If (res) {
      ;  MySendStr .= name . "(" . idx . ")"
      ;  break
      ;}
      If (idx = Cod) {
        MySendStr .= name . "(" . idx . ")"
        break
      }
    }
    A_Clipboard := MySendStr
    SendInput("{Ctrl down}v{Ctrl up}")
  }

}
  
}


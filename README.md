# getElementsByAttributes
vb网页过滤元素的函数，超级好用

## 请将下面的代码放到一个bas模块中
```
Option Explicit

Public Declare Function SafeArrayGetDim Lib "oleaut32.dll " (ByRef saArray() As Any) As Long

'名称：getElementsByAttributes
'功能：根据一个或多个条件对dom对象所有元素进行过滤得到目标元素
'参数：WebBrowser1，WebBrowser类型，要处理的webbrowser
'      strAttributes，string型，内容为属性列表，多项的话用逗号隔开，
'返回：如果有匹配到的结果那么返回的就是html元素对象数组，用户需要执行判断使用
'范例：getElementsByAttributes(WebBrowser1,"id='kw'")(0).value="vb" '设置百度搜索框内容为vb
'      getElementsByAttributes(WebBrowser1,"value='百度一下',type='submit'")(0).click '点击百度的搜索按钮
'      getElementsByAttributes(WebBrowser1,"tagname='input',value^='百度'")(0).click '得到文本开头为“百度”的按钮并执行点击
'作者：sysdzw
'日期：23:53 2017-1-17
Public Function getElementsByAttributes(WebBrowser1 As Object, ByVal strAttributes As String) As Variant
    Dim vTag As Object
    Dim i&, strTiaojians$, isElementOk As Boolean, intElement%, vrt()
    Dim reg As Object
    Dim matchs As Object

    Set reg = CreateObject("vbscript.regExp")
    reg.Global = True
    reg.IgnoreCase = True
    reg.MultiLine = True
    reg.Pattern = "([a-z\dA-Z-_.]+)([!=<>^$*|~]+)(['""]?)([^,]*)\3"
    Set matchs = reg.Execute(strAttributes)

    For Each vTag In WebBrowser1.Document.All
        isElementOk = True
        If LCase(vTag.tagname) = "input" Then
            Dim aa
            aa = 1
        End If
        For i = 0 To matchs.Count - 1
            If LCase(matchs(i).SubMatches(0)) = "tagname" Then
                If Not isConditionOk(LCase(vTag.tagname), matchs(i).SubMatches(1), LCase(matchs(i).SubMatches(3))) Then
                    isElementOk = False
                    Exit For
                End If
            ElseIf LCase(matchs(i).SubMatches(0)) = "innerhtml" Then
                If Not isConditionOk(LCase(vTag.innerhtml), matchs(i).SubMatches(1), LCase(matchs(i).SubMatches(3))) Then
                    isElementOk = False
                    Exit For
                End If
            ElseIf LCase(matchs(i).SubMatches(0)) = "innertext" Then
                If Not isConditionOk(LCase(vTag.innertext), matchs(i).SubMatches(1), LCase(matchs(i).SubMatches(3))) Then
                    isElementOk = False
                    Exit For
                End If
            ElseIf IsNull(vTag.getattribute(matchs(i).SubMatches(0))) Then
                isElementOk = False
                Exit For
            ElseIf Not isConditionOk(vTag.getattribute(matchs(i).SubMatches(0)), matchs(i).SubMatches(1), matchs(i).SubMatches(3)) Then
                isElementOk = False
                Exit For
            End If
        Next

        If isElementOk Then
            ReDim Preserve vrt(intElement)
            Set vrt(intElement) = vTag
            intElement = intElement + 1
        End If
    Next
    If SafeArrayGetDim(vrt) = 0 Then
        getElementsByAttributes = Split("", "")
    Else
        getElementsByAttributes = vrt
    End If
End Function
'简略的调用方法
Public Function g(WebBrowser1 As Object, ByVal strAttributes As String) As Variant
    g = getElementsByAttributes(WebBrowser1, strAttributes)
End Function
'根据运算符检查条件是否符合
Private Function isConditionOk(ByVal strTagValue$, ByVal strCondition$, ByVal strValueForCheck$) As Boolean
    If strCondition = "=" Then
        isConditionOk = (strTagValue = strValueForCheck)
    ElseIf strCondition = "!=" Or strCondition = "<>" Then
        isConditionOk = (strTagValue <> strValueForCheck)
    ElseIf strCondition = "^=" Then '选取开头为strValueForCheck的
        isConditionOk = (Left(strTagValue, Len(strValueForCheck)) = strValueForCheck)
    ElseIf strCondition = "$=" Then '选取末尾为strValueForCheck的
        isConditionOk = (Right(strTagValue, Len(strValueForCheck)) = strValueForCheck)
    ElseIf strCondition = "*=" Then '选取包含strValueForCheck的
        isConditionOk = (InStr(strTagValue, strValueForCheck) > 0)
    ElseIf strCondition = "|=" Then '选取值为strValueForCheck或者值为strValueForCheck前缀的，即strValueForCheck后面加个-
        isConditionOk = (InStr(strTagValue, strValueForCheck) > 0 Or InStr(strTagValue, strValueForCheck & "-") > 0)
    ElseIf strCondition = "~=" Then '选取属性值用空格分隔的值中包含给定值的元素
        isConditionOk = (InStr(" " & strTagValue & " ", " " & strValueForCheck & " ") > 0)
    End If
End Function
```

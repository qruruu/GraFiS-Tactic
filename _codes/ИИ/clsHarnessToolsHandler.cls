Attribute VB_Name = "clsHarnessToolsHandler"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
' Class: clsHarnessToolsHandler - Tools registry and dispatcher for the LLM agent
Option Explicit

Private tools As Collection

Public Sub Init()
    Set tools = New Collection

    Dim toolInstance As clsToolBase

    ' Получение модели пожара
    Set toolInstance = New clsTool_GetFireModel
    tools.Add toolInstance, toolInstance.name

    ' Отрисовка модели пожара
    Set toolInstance = New clsTool_ApplyFireModel
    tools.Add toolInstance, toolInstance.name

    ' Write text file
    Set toolInstance = New clsTool_WriteTextFile
    tools.Add toolInstance, toolInstance.name

    ' Read text file
    Set toolInstance = New clsTool_ReadTextFile
    tools.Add toolInstance, toolInstance.name

    ' Перечень мастеров трафаретов
    Set toolInstance = New clsTool_ListMasters
    tools.Add toolInstance, toolInstance.name

    ' Перечень прототипов мастеров
    Set toolInstance = New clsTool_GetPrototype
    tools.Add toolInstance, toolInstance.name

    ' Перечень фигур
    Set toolInstance = New clsTool_ShapesList
    tools.Add toolInstance, toolInstance.name

    ' Размеры страницы
    Set toolInstance = New clsTool_PageSize
    tools.Add toolInstance, toolInstance.name

End Sub

Public Function GetToolsJson() As String
    If tools Is Nothing Then
        GetToolsJson = "[]"
        Exit Function
    End If

    If tools.Count = 0 Then
        GetToolsJson = "[]"
        Exit Function
    End If

    Dim i As Long
    Dim result As String
    Dim tool As clsToolBase

    result = "["
    For i = 1 To tools.Count
        Set tool = tools.item(i)
        result = result & "{" & _
            """type"":""function""," & _
            """function"":{" & _
                """name"":" & JsonString(tool.name) & "," & _
                """description"":" & JsonString(tool.Description) & "," & _
                """parameters"":" & tool.Schema & _
            "}" & _
        "}"
        If i < tools.Count Then result = result & ","
    Next i

    GetToolsJson = result & "]"
End Function

Public Function Execute(ByVal toolName As String, ByVal argsJson As String) As String
    If tools Is Nothing Then
        Execute = "{""error"":" & JsonString("Tool registry is not initialized") & "}"
        Exit Function
    End If

    On Error GoTo ErrHandler

    Dim tool As clsToolBase
    Set tool = tools.item(toolName)
    Execute = tool.Execute(argsJson)
    Exit Function

ErrHandler:
    Execute = "{""error"":" & JsonString("Tool not found: " & toolName) & "}"
End Function

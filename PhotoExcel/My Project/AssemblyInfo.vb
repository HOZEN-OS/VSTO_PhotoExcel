Imports System
Imports System.Reflection
Imports System.Runtime.InteropServices
Imports System.Security

' アセンブリに関する一般情報は以下の属性セットを通して制御されます。
' アセンブリに関連付けられている情報を変更するには、
' これらの属性値を変更してください。

' アセンブリ属性の値を確認します

<Assembly: AssemblyTitle("PhotoExcel")> 
<Assembly: AssemblyDescription("")> 
<Assembly: AssemblyCompany("")> 
<Assembly: AssemblyProduct("PhotoExcel")> 
<Assembly: AssemblyCopyright("Copyright ©  2021")> 
<Assembly: AssemblyTrademark("")> 

' ComVisible を false に設定すると、その型はこのアセンブリ内で COM コンポーネントには
' 見えなくなります。このアセンブリ内で COM から型にアクセスする必要がある場合は、
' その型の ComVisible 属性を true に設定してください。
<Assembly: ComVisible(False)>

'このプロジェクトが COM に公開される場合、次の GUID が typelib の ID になります
<Assembly: Guid("969615bf-571e-4bd9-9a4f-e0f5d4d87769")> 

' アセンブリのバージョン情報は次の 4 つの値で構成されています:
'
'      メジャー バージョン
'      マイナー バージョン
'      ビルド番号
'      リビジョン
'
' すべての値を指定するか、以下のように '*' を使ってビルドおよびリビジョン番号を
' 既定値にすることができます:
' <Assembly: AssemblyVersion("1.0.*")> 

<Assembly: AssemblyVersion("1.0.0.0")> 
<Assembly: AssemblyFileVersion("1.0.0.0")> 

Friend Module DesignTimeConstants
    Public Const RibbonTypeSerializer As String = "Microsoft.VisualStudio.Tools.Office.Ribbon.Serialization.RibbonTypeCodeDomSerializer, Microsoft.VisualStudio.Tools.Office.Designer, Version=10.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a"
    Public Const RibbonBaseTypeSerializer As String = "System.ComponentModel.Design.Serialization.TypeCodeDomSerializer, System.Design"
    Public Const RibbonDesigner As String = "Microsoft.VisualStudio.Tools.Office.Ribbon.Design.RibbonDesigner, Microsoft.VisualStudio.Tools.Office.Designer, Version=10.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a"
End Module

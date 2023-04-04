'#Reference {59514134-1A57-4156-AD38-7A5B8C7A0055}#6.0#0#C:\Program Files (x86)\Common Files\Kofax\Components\Kofax.Transformation.Extensions.tlb#Kofax_Transformation_Extensions#Kofax_Transformation_Extensions
'#Reference {420B2830-E718-11CF-893D-00A0C9054228}#1.0#0#C:\Windows\SysWOW64\scrrun.dll#Microsoft Scripting Runtime#Scripting
'#Language "WWB-COM"
Option Explicit

' Required reference: C:\Program Files (x86)\Common Files\Kofax\Components\Kofax.Transformation.Extensions.tlb
' Required reference: C:\Windows\SysWOW64\scrrun.dll
Public m_ProjectFolder As String
Public m_ManifestActivator As Object
Private m_AssemblyResolver As Kofax_Transformation_Extensions.AssemblyResolver

Public Function ProjectFolder() As String
   If m_ProjectFolder="" Then
      Dim fso As New FileSystemObject
      m_ProjectFolder = fso.GetFile(Project.FileName).ParentFolder.Path
   End If
   ProjectFolder=m_ProjectFolder
End Function


Public Sub InitializeAssemblyResolver()
   ' Prepare assembly resolver to load assemblies from project subfolder
   ' Note that if the dll is updated it requires restarting Transformation Designer to load the new version.

   If m_AssemblyResolver Is Nothing Then
      Set m_AssemblyResolver = New Kofax_Transformation_Extensions.AssemblyResolver
      m_AssemblyResolver.Init(ProjectFolder() & "\Custom\dlls", Project.ScriptExecutionMode = CscScriptModeServerDesign)

      Set m_ManifestActivator = CreateObject("Microsoft.Windows.ActCtx")
      m_ManifestActivator.manifest = ProjectFolder() & "\Custom\dlls\application.manifest"
   End If
End Sub

Public Function CreateExampleClass() As Object
   InitializeAssemblyResolver()
   Set CreateExampleClass = m_ManifestActivator.CreateObject("RegFreeTransformationDll.ExampleClass")
End Function


Private Sub Document_AfterExtract(ByVal pXDoc As CASCADELib.CscXDocument)
   Dim ExampleClass As Object
   Set ExampleClass=CreateExampleClass()
   Debug.Print ExampleClass.HelloWorld("Project AfterExtract")
End Sub


Private Sub Application_DeinitializeScript()
   If Not m_AssemblyResolver Is Nothing Then m_AssemblyResolver.Shutdown()
End Sub

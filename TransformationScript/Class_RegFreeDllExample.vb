Option Explicit

' Class script: RegFreeDllExample

Private Sub SL_FromDll_LocateAlternatives(ByVal pXDoc As CASCADELib.CscXDocument, ByVal pLocator As CASCADELib.CscXDocField)
   Dim ExampleClass As Object
   Set ExampleClass=CreateExampleClass()

   Dim List As Object
   Set List=ExampleClass.ReturnComplexType()

   Dim ComplexItem As Object
   For Each ComplexItem In List
      Dim Alt As CscXDocFieldAlternative
      Set Alt = pLocator.Alternatives.Create()
      Alt.Confidence=ComplexItem.Confidence
      Alt.Text=ComplexItem.Text
   Next
End Sub

Private Sub SL_FromDll2_LocateAlternatives(ByVal pXDoc As CASCADELib.CscXDocument, ByVal pLocator As CASCADELib.CscXDocField)
   Dim ExampleClass As Object
   Set ExampleClass=CreateExampleClass()

   Dim Alt As CscXDocFieldAlternative
   Set Alt = pLocator.Alternatives.Create()
   Alt.Confidence=1
   Alt.Text=ExampleClass.HelloWorld("SL_FromDll2_LocateAlternatives")
End Sub

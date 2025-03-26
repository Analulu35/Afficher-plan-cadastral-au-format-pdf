Attribute VB_Name = "Cadastre"
#If Win64 Then
Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Declare PtrSafe Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long

#Else
Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long
#End If

Sub AfficheParcelle()
ParcelleCadastre "MERIGNAC", "033", "000", "AT", "146"
End Sub
Public Function ParcelleCadastre(Cne As String, Dep As String, Prefixe As String, Section As String, _
                                 nop As String) As Long
'-------------------------------------------------------------------------------
' Ouvre Google Maps sur le calcul du trajet des adresses passées en arguments.
'-------------------------------------------------------------------------------
Dim objBrowser As New CDPBrowser
Dim listeTrajets As Collection
Dim listeDep As Collection
Dim listeSection As Collection
Dim trajet As CDPElement
Dim sec As CDPElement
Dim result
    On Error Resume Next
   
  
      'on lance MS EDGE
    objBrowser.start "edge", cleanActive:=True, reAttach:=True
    ' on peut cacher le navigateur
    'Optional CacherGoogleMaps As Boolean = False
    If CacherGoogleMaps Then objBrowser.hide
    objBrowser.navigate ("https://www.cadastre.gouv.fr/scpc/rechercherParReferenceCadastrale.do")
    objBrowser.wait till:="interactive"
    Sleep 1000
   
   
    'envoi de l'adresse de départ dans sa zone de saisie dans la page
    objBrowser.getElementByXPath("//input[@name='ville']").value = Cne
    Sleep 1000
  
    
    Set listeDep = objBrowser.getElementsByXPath("//select[@name='codeDepartement']")
    'Sleep 2000
    
    For Each trajet In listeDep
    
       trajet.value = Dep
       trajet.fireEvent "change"
       Sleep 1000
       Exit For
    Next
    
    

    objBrowser.getElementByXPath("//input[@name='prefixeParcelle']").value = Prefixe
    objBrowser.getElementByXPath("//input[@name='sectionLibelle']").value = Section
    objBrowser.getElementByXPath("//input[@name='numeroParcelle']").value = nop
    Sleep 1000
    objBrowser.getElementByXPath("//input[@title='Rechercher']").submit
    Sleep 1000
    objBrowser.activate
     CreateObject("WScript.Shell").sendKeys "{ENTER}"
     Sleep 1000
       CreateObject("WScript.Shell").sendKeys "{TAB 11}"
       CreateObject("WScript.Shell").sendKeys "{ENTER}"
    'objBrowser.getElementByXPath("//a[imprimer']").click
          'Get Window ID for IE so we can set it as activate window
     'HWNDSrc = objIE.HWND
    
        'Set IEDoc1 = IE.document
    Dim objBrowser1 As New CDPBrowser
    Set Fen = objBrowser1.getTab
    
    Sleep 1000
   'Set Fen = objBrowser1.getElementByXPath("//body")
   'Set Form2 = objBrowser1.getElementByID("title_simples_impression")
   'Form2.click
   'a id="menu_advanced"
   Set Fen1 = objBrowser1.getElementByXPath("//a[@id='menu_advanced']")
   'Set Fen1 = objBrowser1.getElementByID("menu_advanced")
   'Fen1.click
   'Fen.getElementByID("menu_advanced").click
   For Each Elt In Fen1
    Debug.Print Fen.Html
       Essai = Elt.getElementByID("menu_advanced")
       
       Sleep 1000
       Exit For
    Next
   
   MsgBox "essai"
  
    'objBrowser.quit
    Set trajet = Nothing: Set listeTrajets = Nothing: Set objBrowser = Nothing
End Function  
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
    Debug.Print Fen.url
    Sleep 1000
    
   'Set Fen = objBrowser1.getElementByXPath("//body")
   Set Form2 = Fen.getElementByID("title_simples_impression")
   'Form2.click
   'a id="menu_advanced"  <h3 id="title_simples_impression" class="clearfix" onclick="if(window.myMenu) myMenu.impression();" onmouseover="switchClass(this);" onmouseout="switchClass(this);"><a>Imprimer</a></h3>
   Set Fen1 = Fen.getElementByXPath("//a[@id='menu_advanced']")
   Fen1.click
   Sleep 500
                                            
                                            
   Set Fen2 = Fen.getElementByID("title_advanced_impression")
   Fen2.click
   'title_advanced_impression
   'Fen.getElementByID("menu_advanced").click
   'tool_avances_impressions_extrait
   Set Fen3 = Fen.getElementByID("tool_avances_impressions_extrait")
   Fen3.click
   Sleep 500
   'a onclick="carte.getWorker().validate();"
   Set Fen4 = Fen.getElementByXPath("//a[@class='action']")
   Fen4.click
   MsgBox "essai"
   
    'objBrowser.quit
    Set trajet = Nothing: Set listeTrajets = Nothing: Set objBrowser = Nothing
End Function
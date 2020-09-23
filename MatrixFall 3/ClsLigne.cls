VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsLigne"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'**********************************************************************************
'*                                                                                *
'* CETTE CLASSE GERE UNE SEULE LIGNE DE PLUSIEURS LETTRES                         *
'*                                                                                *
'* Thomas John (thomas.john@swing.be)                                             *
'*                                                                                *
'**********************************************************************************
'
'classe gérant une lettre
Dim Clsl() As ClsLettre
'
'nombre de lettres chargées
Dim NbrLettres As Long
'
'numéro de la lettre courante
Dim CourLettre As Integer
'
'coordonnées
Dim XSz As Single
Dim YSz As Single
'
'dimensions
Dim HauteurSz As Long
Dim LargeurSz As Long
'
Dim i As Integer
'
'vitesse
Dim VitesseSz As Long
Dim VitesseCnt As Long
'
'temps que va mettre la couleur de la lettre à virer au vert
Dim TransitionSz1 As Long
'
'temps pendant lequel la couleur de la lettre ne changera pas
Dim TransitionSz2 As Long
'
'temps que va mettre la lettre à s'estomper
Dim TransitionSz3 As Long
'
'etat de la ligne
Dim EtatSz As Single '1 = active ; 0 = inactive
'
'flou des lettres
Dim FlouSz As Long
'
'ancienne lettre
Dim AncLettre As Long
'
'
'*********************************
'* PROPRIETES
'*********************************
'
'FLOU DES LETTRES
Public Property Let Flou(Valeur As Long)
    '
    If Valeur = 1 Then
        '
        FlouSz = 26
        '
    Else
        '
        FlouSz = 0
        '
    End If
    '
End Property
'
Public Property Get Flou() As Long
    '
    Flou = FlouSz
    '
End Property
'
'ETAT DE LA LIGNE
Public Property Get Etat() As Single
    '
    Etat = EtatSz
    '
End Property
'
'TRANSITION 1
Public Property Let Transition1(Valeur As Single)
    '
    TransitionSz1 = Valeur
    '
End Property
'
Public Property Get Transition1() As Single
    '
    Transition1 = TransitionSz1
    '
End Property
'
'TRANSITION 2
Public Property Let Transition2(Valeur As Single)
    '
    TransitionSz2 = Valeur
    '
End Property
'
Public Property Get Transition2() As Single
    '
    Transition2 = TransitionSz2
    '
End Property
'
'TRANSITION 3
Public Property Let Transition3(Valeur As Single)
    '
    TransitionSz3 = Valeur
    '
End Property
'
Public Property Get Transition3() As Single
    '
    Transition3 = TransitionSz3
    '
End Property
'
'VITESSE
Public Property Let Vitesse(Valeur As Single)
    '
    VitesseSz = Valeur
    '
End Property
'
Public Property Get Vitesse() As Single
    '
    Vitesse = VitesseSz
    '
End Property
'
'COORDONNEE X
Public Property Let X(Valeur As Single)
    '
    XSz = Valeur
    '
End Property
'
Public Property Get X() As Single
    '
    X = XSz
    '
End Property
'
'HAUTEUR
Public Property Let Hauteur(Valeur As Long)
    '
    HauteurSz = Valeur
    '
End Property
'
Public Property Get Hauteur() As Long
    '
    Hauteur = HauteurSz
    '
End Property
'
'LARGEUR
Public Property Let Largeur(Valeur As Long)
    '
    LargeurSz = Valeur
    '
End Property
'
Public Property Get Largeur() As Long
    '
    Largeur = LargeurSz
    '
End Property
'
'*********************************
'* FIN DES PROPRIETES
'*********************************
'
'
Private Sub Class_Initialize()
    '
    'on initialise certaines variables
    ReDim Clsl(0 To 0)
    '
    NbrLettres = 0
    YSz = 0
    EtatSz = 1
    '
End Sub
'
'AFFICHAGE DES LETTRES LES UNES APRES LES AUTRES
Public Sub Afficher()
    '
    'on ne crée pas de nouvelle lettre temps que le temps (VitesseSz) n'a pas été atteint
    If VitesseCnt = VitesseSz Then
        '
        VitesseCnt = 0
        '
        NbrLettres = NbrLettres + 1
        YSz = (NbrLettres * HauteurSz) - HauteurSz
        '
        'on vérifie que la coordonné y ne dépasse pas la taille de l'affichage choisi
        If YSz < DimH Then
            '
            'on crée une nouvelle lettre
            ReDim Preserve Clsl(0 To NbrLettres)
            '
            Set Clsl(NbrLettres) = New ClsLettre
            '
            Clsl(NbrLettres).Hauteur = HauteurSz
            Clsl(NbrLettres).Largeur = LargeurSz
            Clsl(NbrLettres).X = XSz
            Clsl(NbrLettres).Y = YSz
            Clsl(NbrLettres).Opacite = 255
            Clsl(NbrLettres).Transition1 = TransitionSz1 'RndNbr(5, 24)
            Clsl(NbrLettres).Transition2 = TransitionSz2 'RndNbr(2, 30)
            Clsl(NbrLettres).Transition3 = TransitionSz3
            'Clsl(NbrLettres).Changement = RndNbr(1, 5)
DET_LETTRE:
            Clsl(NbrLettres).Lettre = RndNbr(1, 26) + FlouSz
            '
            'on vérifie que la lettre spécifiée ne soit pas la même que la précédente sinon on recommence
            If Clsl(NbrLettres).Lettre = AncLettre + FlouSz Then GoTo DET_LETTRE
            '
        End If
        '
    Else
        '
        VitesseCnt = VitesseCnt + 1
        '
    End If
    '
    'on continue à afficher les lettres
    For i = 1 To UBound(Clsl)
        '
        'si l'opacité de la lettre est > 0 alors on affiche cette lettre
        If Clsl(i).Opacite > 0 Then
            '
            Clsl(i).AfficherTemps
            '
        End If
        '
    Next
    '
    'on vérifie si l'opacité de la dernière lettre est <= 5
    If UBound(Clsl) = 0 Then Exit Sub
    '
    If Clsl(UBound(Clsl)).Opacite = 0 Then
        '
        'oui, donc on change d'état
        'EtatSz = 0
        ReDim Clsl(0 To 0)
        '
        NbrLettres = 0
        YSz = 0
        EtatSz = 1
        '
    End If
    '
End Sub

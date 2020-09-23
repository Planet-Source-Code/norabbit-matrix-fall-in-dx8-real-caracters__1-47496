VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsLettre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'**********************************************************************************
'*                                                                                *
'* CETTE CLASSE GERE UNE SEULE LETTRE                                             *
'*                                                                                *
'* Thomas John (thomas.john@swing.be)                                             *
'*                                                                                *
'**********************************************************************************
'
'structure vertex pour la fonte matrix
Dim vertMatrix(3) As TLVERTEX
'
'opacité de la lettre
Dim OpaciteSz As Long
'
'numéro de la lettre
Dim NumLettreSz As Long
'
'coordonnées de la lettre
Dim XSz As Single
Dim YSz As Single
'
'dimensions
Dim HauteurSz As Long
Dim LargeurSz As Long
'
'couleur de la lettre
Dim RSz As Long
Dim VSz As Long
Dim BSz As Long
'
'temps de transition 0 (du noir au blanc)
Dim TransTemps0 As Long
'temps écoulé
Dim TempsSz0 As Long
'
'temps de transition 1 (caractère lumineux)
Dim TransTemps1 As Long
'temps écoulé
Dim TempsSz1 As Long
'
'temps de transition 2
Dim TransTemps2 As Long
'temps écoulé 2
Dim TempsSz2 As Long
'
'temps de transition 3 (opacité de la lettre s'estompe)
Dim TransTemps3 As Long
'temps écoulé
Dim TempsSz3 As Long
'
'la lettre peut elle changer dans le temps
Dim ChangementSz As Long
'
'
'*********************************
'* PROPRIETES
'*********************************
'
'VITESSE A LAQUELLE L'OPACITE DE LA LETTRE CHANGE
Public Property Let Transition3(Valeur As Long)
    '
    TransTemps3 = Valeur
    '
End Property
'
Public Property Get Transition3() As Long
    '
    Transition3 = TransTemps3
    '
End Property
'
'TEMPS DE LA 2EME TRANSITION
Public Property Let Transition2(Valeur As Long)
    '
    TransTemps2 = Valeur
    '
End Property
'
Public Property Get Transition2() As Long
    '
    Transition2 = TransTemps2
    '
End Property
'
'TEMPS DE VIE DU CARACTERE LUMINEUX
Public Property Let Transition1(Valeur As Long)
    '
    TransTemps1 = Valeur
    '
End Property
'
Public Property Get Transition1() As Long
    '
    Transition1 = TransTemps1
    '
End Property
'
'LA LETTRE PEUT ELLE CHANGER DANS LE TEMPS
Public Property Let Changement(Valeur As Long)
    '
    ChangementSz = Valeur
    '
End Property
'
Public Property Get Changement() As Long
    '
    Changement = ChangementSz
    '
End Property
'
'NUMERO DE LA LETTRE
Public Property Let Lettre(Valeur As Long)
    '
    NumLettreSz = Valeur
    '
End Property
'
Public Property Get Lettre() As Long
    '
    Lettre = NumLettreSz
    '
End Property
'
'OPACITE
Public Property Let Opacite(Valeur As Long)
    '
    OpaciteSz = Valeur
    '
End Property
'
Public Property Get Opacite() As Long
    '
    Opacite = (RSz + VSz + BSz) \ 3
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
'COORDONNEE Y
Public Property Let Y(Valeur As Single)
    '
    YSz = Valeur
    '
End Property
'
Public Property Get Y() As Single
    '
    Y = YSz
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
'procédure permettan d'affiche la fonte choisie (via une texture)
Private Sub AfficherFonte(NuméroFonte As Long, Optional CouleurSz As Long = vbWhite)
    '
    Dim CharX As Integer
    Dim CharY As Integer
    Dim cTmp As Long
    '
    'on calcule les coordonnées
    cTmp = (NuméroFonte - 1) \ 13
    CharY = 20 * cTmp
    cTmp = (NuméroFonte - 1) Mod 13
    CharX = 13 * cTmp
    '
    'on génère les points
    vertMatrix(0) = CreateTLVertex(XSz, YSz, 0, 1, CouleurSz, 0, (1 / 256) * CharX, (1 / 128) * CharY)
    vertMatrix(1) = CreateTLVertex(XSz + LargeurSz, YSz, 0, 1, CouleurSz, 0, ((1 / 256) * CharX) + ((1 / 256) * 13), (1 / 128) * CharY)
    vertMatrix(2) = CreateTLVertex(XSz, YSz + HauteurSz, 0, 1, CouleurSz, 0, (1 / 256) * CharX, ((1 / 128) * CharY) + ((1 / 128) * 20))
    vertMatrix(3) = CreateTLVertex(XSz + LargeurSz, YSz + HauteurSz, 0, 1, CouleurSz, 0, ((1 / 256) * CharX) + ((1 / 256) * 13), ((1 / 128) * CharY) + ((1 / 128) * 20))
    
    'on dessine l'image
    D3DDevice.SetTexture 0, MatrixTex
    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, vertMatrix(0), Len(vertMatrix(0))
    '
End Sub
'
'procédure qui va s'occuper d'afficher la lettre dans le temps
Public Sub AfficherTemps()
    '
    'transition du noir vers le blanc
    If TempsSz0 < TransTemps0 Then
        '
        'on change la couleur de la lettre en fonction du temps que cela doit prendre
        RSz = RSz - ((RSz - 255) / (TransTemps0 - TempsSz0))
        VSz = VSz - ((VSz - 255) / (TransTemps0 - TempsSz0))
        BSz = BSz - ((BSz - 255) / (TransTemps0 - TempsSz0))
        '
        AfficherFonte NumLettreSz, RGB(RSz, VSz, BSz)
        '
        TempsSz0 = TempsSz0 + 1
        '
        Exit Sub
        '
    End If
    '
    'transition de la couleur blanche vers la couleur verte
    If TempsSz1 < TransTemps1 Then
        '
        'on change la couleur de la lettre en fonction du temps que cela doit prendre
        RSz = RSz - ((RSz - 68) / (TransTemps1 - TempsSz1))
        VSz = VSz - ((VSz - 129) / (TransTemps1 - TempsSz1))
        BSz = BSz - ((BSz - 82) / (TransTemps1 - TempsSz1))
        '
        AfficherFonte NumLettreSz, RGB(RSz, VSz, BSz)
        '
        TempsSz1 = TempsSz1 + 1
        '
        Exit Sub
        '
    End If
    '
    'on effectue la transition suivante
    If TempsSz2 < TransTemps2 Then
        '
        'ici, la couleur reste la même
        AfficherFonte NumLettreSz, RGB(RSz, VSz, BSz)
        '
        TempsSz2 = TempsSz2 + 1
        '
        Exit Sub
        '
    End If
    '
    'on effectue la transition suivante
    If TempsSz3 < TransTemps3 Then
        '
        'ici, la couleur s'obscurci
        RSz = RSz - (RSz / (TransTemps3 - TempsSz3))
        VSz = VSz - (VSz / (TransTemps3 - TempsSz3))
        BSz = BSz - (BSz / (TransTemps3 - TempsSz3))
        '
        AfficherFonte NumLettreSz, RGB(RSz, VSz, BSz)
        '
        TempsSz3 = TempsSz3 + 1
        '
        Exit Sub
        '
    End If
    '
    'si la lettre peut changer dans le temps, on la change
    'If ChangementSz = 1 Then NumLettreSz = RndNbr(1, 52)
    '
End Sub

Private Sub Class_Initialize()
    '
    RSz = 255
    VSz = 255
    BSz = 255
    '
End Sub

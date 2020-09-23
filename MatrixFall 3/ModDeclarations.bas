Attribute VB_Name = "ModDeclarations"
'*************************************************************************
'*                                                                       *
'* MODULE DX8                                                            *
'*                                                                       *
'* DECLARATION DES VARIABLES PRINCIPALES + FONCTIONS                     *
'*                                                                       *
'* traduit et complété par Thomas John (thomas.john@swing.be)            *
'*                                                                       *
'* source : http://216.5.163.53/DirectX4VB (DirectX 4 VB, Jack Hoxley)   *
'*                                                                       *
'*************************************************************************
'
'l'objet principal
Public Dx As DirectX8
'
'cet objet contrôle tout ce qui est 3D
Public D3D As Direct3D8
'
'cet objet représente le "hardware" (la carte graphique) utilisé pour le rendu
Public D3DDevice As Direct3DDevice8
'
'une "librairie d'aide"
'D3DX8 est une classe d'aide qui contient une multitude de fonctions destinées à faciliter la programmation en DX8
Public D3DX As D3DX8
'
'variable servant à détecter si le programme tourne ou pas
Public bRunning As Boolean
'
'ceci est une description d'un format flexible pour un vertex 2D
Public Const TL_FVF = D3DFVF_XYZRHW Or D3DFVF_TEX1 Or D3DFVF_DIFFUSE Or D3DFVF_SPECULAR
'
'cette structure représente un vertex (identique à la structure "D3DTLVERTEX" pour Dx7)
Public Type TLVERTEX
    X As Single
    Y As Single
    z As Single
    rhw As Single
    color As Long
    Specular As Long
    tu As Single
    tv As Single
End Type
'
'fonte
Public MainFont As D3DXFont
Public MainFontDesc As IFont
Public TextRect As RECT
Public fnt As New StdFont
'
'Pi
Public Const pi As Single = 3.14159265358979
'
Public matWorld As D3DMATRIX '//How the vertices are positioned
'où la caméra se trouve et vers où pointe-t-elle
Public matView As D3DMATRIX
'comment la caméra projecte le monde 3D sur un écran (surface) 2D
Public matProj As D3DMATRIX
'
'calcul du fps (images par seconde)
Public FPS_NbrFps As Long
Public FPS_Temps As Long
Public FPS_NbrImg As Long
'
'dimensions de l'affichage
Public DimH As Long
Public DimL As Long
'
'test
Public NbrSz As Long
'
'classe représentant une ligne
Public Clsl() As ClsLigne
'
'texture de fonte matrix
Public MatrixTex As Direct3DTexture8
'
'pause
Public PauseSz As Boolean
'
'
'*******************************************************************
'* Initialise
'*******************************************************************
'
Public Function Initialise(FrmObjet As Form, DimLargeur As Long, DimHauteur As Long) As Boolean
    '
    'On Error GoTo ErrHandler:
    '
    'décrit notre mode d'affichage
    Dim DispMode As D3DDISPLAYMODE
    '
    'décrit notre mode de vue
    Dim D3DWindow As D3DPRESENT_PARAMETERS
    '
    'pour les filtreurs de texture
    Dim Caps As D3DCAPS8 '//For Texture Filters
    '
    'on crée notre objet principal
    Set Dx = New DirectX8
    '
    'on crée l'interface Direct3D via notre objet principal
    Set D3D = Dx.Direct3DCreate()
    '
    'on crée notre librairie d'aide
    Set D3DX = New D3DX8
    '
    'DispMode.Format = D3DFMT_X8R8G8B8
    DispMode.Format = D3DFMT_R5G6B5 'si ce mode ne fonctionne pas, utilisez celui juste au-dessus
    DispMode.Width = DimLargeur
    DispMode.Height = DimHauteur
    '
    DimL = DimLargeur
    DimH = DimHauteur
    '
    D3DWindow.SwapEffect = D3DSWAPEFFECT_FLIP
    D3DWindow.BackBufferCount = 1 '1 BackBuffer
    D3DWindow.BackBufferFormat = DispMode.Format
    D3DWindow.BackBufferWidth = DispMode.Width
    D3DWindow.BackBufferHeight = DispMode.Height
    D3DWindow.hDeviceWindow = FrmObjet.hWnd
    D3DWindow.EnableAutoDepthStencil = 1
    '
    If D3D.CheckDeviceFormat(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, DispMode.Format, D3DUSAGE_DEPTHSTENCIL, D3DRTYPE_SURFACE, D3DFMT_D16) = D3D_OK Then
        '
        'on peut utiliser un Z-Buffer de 16-bit
        D3DWindow.AutoDepthStencilFormat = D3DFMT_D16
        '
    Else
        '
        'à mettre ici : procédure de détection des différents modes d'affichage
        '
    End If
    '
    'cette ligne crée un "device" qui utilise la carte graphique ("hardware") pour effectuer les calculs si possible,
    'ou le processeur ("software") et utilise comme objet de réception notre feuille principale
    Set D3DDevice = D3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, FrmObjet.hWnd, D3DCREATE_SOFTWARE_VERTEXPROCESSING, D3DWindow)
    '
    'on demande au vertex shader d'utiliser notre format de vertex
    D3DDevice.SetVertexShader TL_FVF
    '
    'nos points (vertices) n'ont pas besoin de lumière, donc on désactive cette option
    D3DDevice.SetRenderState D3DRS_LIGHTING, False
    '
    'D3DDevice.SetRenderState D3DRS_CULLMODE, D3DCULL_NONE
    '
    'déclarations utiles pour le rendu de textures transparantes (par couleur)
    D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
    D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
    D3DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, True
    '
    'filtrage de texture : donne un meilleur résultat lors d'un redimensionnement d'une texture
    D3DDevice.SetTextureStageState 0, D3DTSS_MAGFILTER, D3DTEXF_LINEAR
    D3DDevice.SetTextureStageState 0, D3DTSS_MINFILTER, D3DTEXF_LINEAR
    '
    'on active notre Z-Buffer
    D3DDevice.SetRenderState D3DRS_ZENABLE, 1
    '
    '
    'la matrice "World"
    D3DXMatrixIdentity matWorld
    D3DDevice.SetTransform D3DTS_WORLD, matWorld
    '
    'la matrice "View"
    D3DXMatrixLookAtLH matView, MakeVector(0, 9, -9), MakeVector(0, 0, 0), MakeVector(0, 1, 0)
    D3DDevice.SetTransform D3DTS_VIEW, matView
    '
    'la matrice de projection
    D3DXMatrixPerspectiveFovLH matProj, pi / 4, 1, 0.1, 500
    D3DDevice.SetTransform D3DTS_PROJECTION, matProj
    '
    '
    'initialisation du rendu du texte
    fnt.Name = "Tahoma"
    fnt.Size = 12
    fnt.Bold = True
    Set MainFontDesc = fnt
    Set MainFont = D3DX.CreateFont(D3DDevice, MainFontDesc.hFont)
    '
    '**************************************
    '** chargement des textures          **
    '**************************************
    '
    Set MatrixTex = D3DX.CreateTextureFromFileEx(D3DDevice, App.Path & "\fontes.png", 256, 128, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_POINT, 0, ByVal 0, ByVal 0)
    '
    '**************************************
    '** fin du chargement des textures   **
    '**************************************
    '
    'si on arrive jusqu'ici, c'est que tout s'est bien passé
    Initialise = True
    '
    'on charge les lignes
    ChargerLignesMatrix
    'ChargerLignesMatrix
    '
    Dim RotateAngle As Single
    Dim matTemp As D3DMATRIX 'contient des données temporaires
    '
    'on montre notre feuille pour être spur
    FrmObjet.Show
    '
    bRunning = True
    '
    Do While bRunning
        '
        'on "rend" (dessine) la scène
        Render
        '
        'on calcule le nombre d'images par secondes
        If GetTickCount() - FPS_Temps >= 100 Then
            '
            FPS_NbrFps = FPS_NbrImg * (1000 / (GetTickCount() - FPS_Temps))
            FPS_Temps = GetTickCount()
            FPS_NbrImg = 0
            '
        End If
        
        FPS_NbrImg = FPS_NbrImg + 1
        '
        'on laisse vb respirer
        DoEvents
        '
    Loop
    '
    'la boucle s'est terminée signifiant que le programme va se fermer
    'il faut avant tout décharger les objets qu'on a chargé précédemment
    '
    On Error Resume Next 'pour être sûr
    '
    Set D3DDevice = Nothing
    Set D3D = Nothing
    Set Dx = Nothing
    '
    'fin
    Exit Function
    '
ErrHandler:
    '
    'un erreur s'est produite
    MsgBox "Error Number Returned: " & Err.Number & " [" & Err.Description & "]"
    '
    Initialise = False
    '
End Function
'
'PROCEDURE DE RENDU DE LA SCENE
Public Sub Render()
    '
    If PauseSz = True Then Exit Sub
    '
    Dim i As Integer
    '
    'on efface la surface et on lui donne la couleur noir
    D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET Or D3DCLEAR_ZBUFFER, 0, 1#, 0
    '
    'on commence le rendu
    D3DDevice.BeginScene
        '
        'tous les appels de rendu doivent être fait entre "BeginScene" et "EndScene"
        '
        'rendu du texte
        TextRect.Top = 1
        TextRect.Left = 1
        TextRect.bottom = 480
        TextRect.Right = 640
        D3DX.DrawText MainFont, &HFFCCCCFF, FPS_NbrFps & "fps", TextRect, DT_LEFT
        '
        'on affiche les lettres
        For i = LBound(Clsl) + 1 To UBound(Clsl)
            '
            Clsl(i).Afficher
            '
        Next
        '
    D3DDevice.EndScene
    '
    'on met à jour l'image à l'écran
    D3DDevice.Present ByVal 0, ByVal 0, 0, ByVal 0
    '
End Sub
'
'fonction permettant de créer un vecteur en une ligne
Public Function MakeVector(X As Single, Y As Single, z As Single) As D3DVECTOR
    '
    MakeVector.X = X
    MakeVector.Y = Y
    MakeVector.z = z
    '
End Function
'
'fontion permettant de "couvrir" (texture) une structure plus simplement
Public Function CreateTLVertex(X As Single, Y As Single, z As Single, rhw As Single, color As Long, Specular As Long, tu As Single, tv As Single) As TLVERTEX
    '//NB: whilst you can pass floating point values for the coordinates (single)
    '       there is little point - Direct3D will just approximate the coordinate by rounding
    '       which may well produce unwanted results....
    CreateTLVertex.X = X
    CreateTLVertex.Y = Y
    CreateTLVertex.z = z
    CreateTLVertex.rhw = rhw
    CreateTLVertex.color = color
    CreateTLVertex.Specular = Specular
    CreateTLVertex.tu = tu
    CreateTLVertex.tv = tv
    '
End Function
'
'fonction renvoyant un nombre aléatoire situé entre un minimum et un maximum
'(merci à raff (VbFrance) pour le code disponible à l'adresse suivante : http://www.vbfrance.com/article.aspx?ID=7209)
'(code légèrement modifié)
Public Function RndNbr(MinSz As Long, MaxSz As Long) As Long

     RndNbr = (Rnd * (MaxSz - MinSz)) + MinSz
                
End Function
'
'cette procédure charge autant de lignes qu'il est nécessaire pour la largeur de l'affichage spécifié
Public Sub ChargerLignesMatrix()
    '
    Dim xTmp As Long
    Dim iTmp As Integer
    Dim lTmp As Long
    '
    Do
        '
        'nouvelle hauteur
        lTmp = RndNbr(10, 25)
        '
        'on charge un nouvelle ligne
        iTmp = ChargerLigne(xTmp, lTmp, lTmp * 0.65, RndNbr(1, 5), RndNbr(14, 22), RndNbr(2, 4), RndNbr(30, 70))
        '
        'on déplace la position X pour la ligne suivante
        xTmp = xTmp + Clsl(iTmp).Largeur
        '
        'on vérifie qu'on ne sort pas de l'écran
        If xTmp >= DimL Then Exit Do
        '
    Loop
    '
End Sub
'
'cette fontion charge une ligne et renvoie son index
Public Function ChargerLigne(XSz As Long, DimHauteurSz As Long, DimLargeurSz As Long, VitesseSz As Long, Trans1 As Long, Trans2 As Long, Trans3 As Long) As Integer
    '
    ReDim Preserve Clsl(LBound(Clsl) To UBound(Clsl) + 1)
    '
    Set Clsl(UBound(Clsl)) = New ClsLigne
    '
    'dimensions des lettres
    Clsl(UBound(Clsl)).Hauteur = DimHauteurSz 'RndNbr(10, 25)
    Clsl(UBound(Clsl)).Largeur = DimLargeurSz 'Clsl(UBound(Clsl)).Hauteur * 0.65
    '
    'position de la ligne
    Clsl(UBound(Clsl)).X = XSz
    '
    'vitesse des lettres
    Clsl(UBound(Clsl)).Vitesse = VitesseSz
    '
    'durée de la transition vers le vert
    Clsl(UBound(Clsl)).Transition1 = Trans1
    '
    'durée de la transition 2
    Clsl(UBound(Clsl)).Transition2 = Trans2
    '
    'durée de la transition vers le noir
    Clsl(UBound(Clsl)).Transition3 = Trans3
    '
    'on détermine le flou des lettres en fonction de leur vitesse
    Clsl(UBound(Clsl)).Flou = Clsl(UBound(Clsl)).Vitesse
    '
    ChargerLigne = UBound(Clsl)
    '
End Function

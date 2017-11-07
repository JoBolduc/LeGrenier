Attribute VB_Name = "modVarglobales"
Option Compare Database
Option Explicit

' Variable de l'application
Global gClient As FicheClient
Global gMenage As InformationMenage
Global Const cnsEnvironnement = "DEV"
Global Const cnsVersion = "2.5.0"
Global Const cnsConnectData = "MS Access;PWD=legrenier_pwd;DATABASE=C:\LeGrenier\BD\Comptoir_Data.accdb"
Global Const cnsConnectTemp = ";DATABASE=C:\LeGrenier\BD\Comptoir_Temp.accdb"
Global Const cnsCheminBDTemp = "C:\LeGrenier\BD\Comptoir_Temp.accdb"
Global Const cnsCheminBDData = "C:\LeGrenier\BD\Comptoir_Data.accdb"
Global Const cnsCheminBDStats = "C:\LeGrenier\BD\Comptoir_Stats.accdb"
Global Const cnsCheminDestRapport = "C:\LeGrenier\Rapports\"
Global Const cnsModeleRapportHebdo = "C:\LeGrenier\Rapports\ModeleRappHebdoMensAnn.xlsx"
Global Const cnsModeleRapportFacturation = "C:\LeGrenier\Rapports\ModeleFacturation.xlsx"
Global Const cnsModeleRapportStats = "C:\LeGrenier\Rapports\ModeleStats.xlsb"

' Variable pour mise en production des rapports
' Le rapport doit être déposé à la racine de l'application
Global Const cnsModeleHebdoMEP = "C:\LeGrenier\ModeleRappHebdoMensAnn.xlsx"
Global Const cnsModeleRapportFacturationMEP = "C:\LeGrenier\ModeleFacturation.xlsx"
Global Const cnsImageGrenier = "C:\LeGrenier\legrenier-logo.jpg"
' Variable des énums
Global geSexe As eSexe
Global geTypMenage As eTypMenage
Global geSourceRevenu As eSourceRevenu
Global geTypLogement As eTypLogement
Global geServiceUtil() As eServiceUtil

' Déclaration des enum pour tous les types de choix clients
Enum eSexe
    Masculin ' 0
    Feminin ' 1
End Enum

Enum eTypMenage
    Monoparentale ' 1
    Biparentale ' 2
    CoupleSansEnfant ' 3
    Celibataire ' 4
End Enum

Enum eSourceRevenu
    Emploi ' 1
    AssEmploi ' 2
    AideSociale ' 3
    AideInvalidite ' 4
    Pension ' 5
    AideEtude ' 6
    SansRevenu ' 7
    Autre ' 8
End Enum

Enum eTypLogement
    Proprietaire ' 1
    LogPrive ' 2
    LogSocial ' 3
    LogBande ' 4
    RefugeUrg ' 5
    HebJeune ' 6
    Rue ' 7
    Famille ' 8
End Enum

Enum eServiceUtil
    AideAlimentaire
    CuisineCollect
    DinerCommun
    Benevolat
End Enum



    

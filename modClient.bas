Attribute VB_Name = "modClient"
Option Compare Database
Option Explicit


Public Type FicheClient
    NoClient As Integer
    Nom As String
    Prenom As String
    Sexe As String
    DateNaissance As Date
    DateInscription As Date
    Adresse As String
    TelephonePrinc As String
    TelephoneSec As String
    Ville As String
    EtudPostSec As Integer
    PremiereNation As Integer
    Immigrant As Integer
    TypeMenage As Integer
    SourceRevenu As Integer
    TypeLogement As Integer
    AideAlim As Integer
    CuisCollect As Integer
    DinerCommun As Integer
    Benevolat As Integer
    Commentaire As String
    Actif As Boolean
End Type

Public Type InformationMenage
    NoClient As Integer
    NoMenage As Integer
    Nom As String
    Prenom As String
    Sexe As String
    DateNaissance As Date
    Lien As String
End Type

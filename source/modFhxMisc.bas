Attribute VB_Name = "Fhx_Misc"
Option Explicit

Public Type fileheader     '16 bytes
    fourcc As Long          '4 bytes
    version As Integer      '2 bytes
    reserved As Integer     '2 bytes
    size As Long            '4 bytes
    offset As Long          '4 bytes
End Type


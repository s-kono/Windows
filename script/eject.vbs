Option Explicit

Dim wmp

Set wmp = CreateObject("WMPlayer.OCX")

wmp.cdromcollection.item(0).eject()
'wmp.cdromcollection.item(1).eject()

WScript.Quit()
